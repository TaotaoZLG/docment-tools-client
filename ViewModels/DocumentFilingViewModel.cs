using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection.Metadata;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Input;
using docment_tools_client.Helpers;
using docment_tools_client.Models;
using Microsoft.Win32;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using SharpCompress.Archives;
using SharpCompress.Common;
using SharpCompress.Readers;
using Tesseract;

namespace docment_tools_client.ViewModels
{
    public class DocumentFilingViewModel : ViewModelBase, IDisposable
    {
        #region 绑定属性
        // ========== 核心修改：重命名压缩包路径为源路径（支持文件/文件夹） ==========
        private string _sourcePath;
        public string SourcePath
        {
            get => _sourcePath;
            set { _sourcePath = value; OnPropertyChanged(); }
        }

        private string _exportFolderPath;
        public string ExportFolderPath
        {
            get => _exportFolderPath;
            set { _exportFolderPath = value; OnPropertyChanged(); }
        }

        private string _excelTemplateFilePath;
        public string ExcelTemplateFilePath
        {
            get => _excelTemplateFilePath;
            set { _excelTemplateFilePath = value; OnPropertyChanged(); }
        }

        private ObservableCollection<FilingRule> _filingRuleList;
        public ObservableCollection<FilingRule> FilingRuleList
        {
            get => _filingRuleList;
            set { _filingRuleList = value; OnPropertyChanged(); }
        }

        private FilingRule _selectedFilingRule;
        public FilingRule SelectedFilingRule
        {
            get => _selectedFilingRule;
            set { _selectedFilingRule = value; OnPropertyChanged(); }
        }

        private ObservableCollection<string> _ocrLanguageList;
        public ObservableCollection<string> OcrLanguageList
        {
            get => _ocrLanguageList;
            set { _ocrLanguageList = value; OnPropertyChanged(); }
        }

        private string _selectedOcrLanguage;
        public string SelectedOcrLanguage
        {
            get => _selectedOcrLanguage;
            set { _selectedOcrLanguage = value; OnPropertyChanged(); }
        }

        private ObservableCollection<UILogItem> _logList;
        public ObservableCollection<UILogItem> LogList
        {
            get => _logList;
            set { _logList = value; OnPropertyChanged(); }
        }

        // 临时归档目录（避免磁盘残留）
        private readonly string _tempArchiveDir;
        // 生成的Excel路径
        private string _excelFilePath;
        // 标记是否已释放资源
        private bool _disposed = false;
        #endregion

        #region 命令定义
        // ========== 核心修改：重命名选择压缩包命令为选择源路径命令 ==========
        public ICommand SelectSourcePathCommand { get; }
        public ICommand ProcessSourceAndArchiveCommand { get; } // 原UnzipAndArchiveCommand
        public ICommand OcrAndGenerateExcelCommand { get; }
        public ICommand SelectExportFolderCommand { get; }
        public ICommand ExportFilesCommand { get; }
        public ICommand SelectExcelTemplateCommand { get; }
        #endregion

        public DocumentFilingViewModel()
        {
            // 初始化日志集合
            LogList = new ObservableCollection<UILogItem>();
            // 初始化临时目录（带GUID避免冲突）
            _tempArchiveDir = Path.Combine(Path.GetTempPath(), $"DocArchive_{Guid.NewGuid():N}");
            Directory.CreateDirectory(_tempArchiveDir);

            // 初始化命令
            SelectSourcePathCommand = new RelayCommand(SelectSourcePath); // 替换原SelectCompressFileCommand
            ProcessSourceAndArchiveCommand = new RelayCommand(ProcessSourceAndArchive); // 替换原UnzipAndArchiveCommand
            OcrAndGenerateExcelCommand = new RelayCommand(OcrAndGenerateExcel);
            SelectExportFolderCommand = new RelayCommand(SelectExportFolder);
            ExportFilesCommand = new RelayCommand(ExportFiles);
            SelectExcelTemplateCommand = new RelayCommand(SelectExcelTemplate);

            // 初始化归档规则（优先级从高到低）
            InitFilingRules();
            // 初始化OCR语言选项
            InitOcrLanguages();

            AddLog("文档归档工具初始化完成");
        }

        #region 初始化方法
        /// <summary>
        /// 初始化归档规则（优先级：身份证号 > 唯一字符串 > 文件名）
        /// </summary>
        private void InitFilingRules()
        {
            FilingRuleList = new ObservableCollection<FilingRule>
            {
                // 规则1（最高优先级）：匹配姓名+18位身份证号
                new FilingRule
                {
                    RuleName = "1. 姓名+18位身份证号（优先）",
                    GetArchiveKey = fileName =>
                    {
                        // 优化身份证正则：严格匹配18位（最后一位支持X/x）
                        var idCardRegex = new Regex(@"[1-9]\d{5}(19|20)\d{2}((0[1-9])|(1[0-2]))(([0-2][1-9])|10|20|30|31)\d{3}[\dXx]");
                        var idCardMatch = idCardRegex.Match(fileName);
                        if (idCardMatch.Success)
                        {
                            // 提取身份证号前的姓名（仅中文，最多提取4个汉字）
                            var nameSegment = fileName.Substring(0, idCardMatch.Index);
                            var nameMatch = Regex.Match(nameSegment, @"[\u4e00-\u9fa5]{1,4}");
                            var name = nameMatch.Success ? nameMatch.Value.Trim() : string.Empty;
                            return string.IsNullOrEmpty(name) ? idCardMatch.Value : $"{name}_{idCardMatch.Value}";
                        }
                        // 匹配不到则降级到规则2
                        return GetUniqueStringKey(fileName);
                    }
                },
                // 规则2：匹配唯一字符串（字母/数字/下划线组合）
                new FilingRule
                {
                    RuleName = "2. 唯一字符串（字母/数字/下划线）",
                    GetArchiveKey = fileName => GetUniqueStringKey(fileName)
                },
                // 规则3（兜底）：仅按文件名（去后缀）
                new FilingRule
                {
                    RuleName = "3. 文件名（去后缀）（兜底）",
                    GetArchiveKey = fileName => Path.GetFileNameWithoutExtension(fileName)
                }
            };
            // 默认选中最高优先级规则
            SelectedFilingRule = FilingRuleList.First();
        }

        /// <summary>
        /// 提取唯一字符串（字母/数字/下划线，长度≥6）
        /// </summary>
        private string GetUniqueStringKey(string fileName)
        {
            // 优化正则：匹配连续的字母/数字/下划线（长度≥6）
            var uniqueStrMatch = Regex.Match(fileName, @"(?<=[^a-zA-Z0-9_])[a-zA-Z0-9_]{6,}(?=[^a-zA-Z0-9_])");
            return uniqueStrMatch.Success ? uniqueStrMatch.Value : Path.GetFileNameWithoutExtension(fileName);
        }

        /// <summary>
        /// 初始化OCR语言选项
        /// </summary>
        private void InitOcrLanguages()
        {
            OcrLanguageList = new ObservableCollection<string> { "简体中文", "英文", "中英双语" };
            SelectedOcrLanguage = "简体中文";
        }
        #endregion

        #region 核心业务逻辑
        /// <summary>
        /// 选择源路径（支持压缩包文件/文件夹）
        /// </summary>
        private void SelectSourcePath()
        {
            // 弹出选择类型对话框，让用户选择是选文件还是文件夹
            var result = MessageBox.Show("请选择要处理的类型：\n【是】- 压缩包文件（zip/rar/7z）\n【否】- 文件夹", "选择处理类型",
                MessageBoxButton.YesNoCancel, MessageBoxImage.Question);

            if (result == MessageBoxResult.Cancel)
                return;

            if (result == MessageBoxResult.Yes)
            {
                // 选择压缩包文件
                var openFileDialog = new OpenFileDialog
                {
                    Title = "选择压缩包文件",
                    Filter = "压缩包文件|*.zip;*.rar;*.7z|所有文件|*.*",
                    RestoreDirectory = true
                };

                if (openFileDialog.ShowDialog() == true)
                {
                    // 校验文件格式是否支持
                    var ext = Path.GetExtension(openFileDialog.FileName).ToLower();
                    if (!new[] { ".zip", ".rar", ".7z" }.Contains(ext))
                    {
                        MessageBox.Show("暂不支持该压缩格式，请选择zip/rar/7z文件！", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }

                    SourcePath = openFileDialog.FileName;
                    AddLog($"已选择压缩包：{Path.GetFileName(SourcePath)}");
                }
            }
            else
            {
                // 选择文件夹
                var folderDialog = new OpenFolderDialog
                {
                    Title = "选择要归档的文件夹",
                    InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
                };

                if (folderDialog.ShowDialog() == true)
                {
                    SourcePath = folderDialog.FolderName;
                    AddLog($"已选择文件夹：{SourcePath}");
                }
            }
        }

        /// <summary>
        /// 选择Excel模板文件
        /// </summary>
        private void SelectExcelTemplate()
        {
            var openFileDialog = new OpenFileDialog
            {
                Title = "选择Excel模板文件",
                Filter = "Excel文件|*.xlsx;*.xls|所有文件|*.*",
                RestoreDirectory = true
            };

            if (openFileDialog.ShowDialog() == true)
            {
                ExcelTemplateFilePath = openFileDialog.FileName;
                AddLog($"已选择Excel模板：{Path.GetFileName(ExcelTemplateFilePath)}");
            }
        }

        /// <summary>
        /// 处理源路径（压缩包解压/文件夹直接遍历）并按规则归档
        /// </summary>
        private void ProcessSourceAndArchive()
        {
            try
            {
                if (string.IsNullOrEmpty(SourcePath) || (!File.Exists(SourcePath) && !Directory.Exists(SourcePath)))
                {
                    MessageBox.Show("请先选择有效的压缩包文件或文件夹！", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                AddLog("开始处理源路径并归档文件...");
                // 清空临时目录（避免残留）
                CleanDirectory(_tempArchiveDir);
                Directory.CreateDirectory(_tempArchiveDir);

                var fileCount = 0;

                // ========== 核心逻辑：区分源路径类型（文件/文件夹） ==========
                if (File.Exists(SourcePath))
                {
                    // 源路径是压缩包文件 → 解压并归档
                    fileCount = ProcessCompressFile(SourcePath);
                }
                else if (Directory.Exists(SourcePath))
                {
                    // 源路径是文件夹 → 直接遍历并归档
                    fileCount = ProcessFolder(SourcePath);
                }

                // 清理空文件夹
                CleanEmptyFolders(_tempArchiveDir);

                AddLog($"处理归档完成！共处理 {fileCount} 个文件，临时目录：{_tempArchiveDir}");
            }
            catch (Exception ex)
            {
                AddLog($"处理归档失败：{ex.Message}", LogLevel.Error);
                MessageBox.Show($"处理失败：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 处理压缩包文件（解压并归档）
        /// </summary>
        private int ProcessCompressFile(string compressFilePath)
        {
            int fileCount = 0;
            // 适配 SharpCompress 0.46.3：使用 ArchiveFactory.Open()
            using (var archive = ArchiveFactory.Open(compressFilePath))
            {
                foreach (var entry in archive.Entries)
                {
                    // 跳过目录和空文件
                    if (entry.IsDirectory || entry.Size == 0)
                        continue;

                    try
                    {
                        // 按选中规则生成归档Key
                        var archiveKey = SelectedFilingRule.GetArchiveKey(entry.Key);
                        var archiveFolder = Path.Combine(_tempArchiveDir, archiveKey);
                        // 创建归档文件夹
                        if (!Directory.Exists(archiveFolder))
                        {
                            Directory.CreateDirectory(archiveFolder);
                            AddLog($"创建归档文件夹：{archiveKey}");
                        }

                        // 适配 SharpCompress 0.46.3：解压文件到归档文件夹
                        var destFilePath = Path.Combine(archiveFolder, Path.GetFileName(entry.Key));
                        // 使用 entry.WriteToFile() 配合 ExtractionOptions
                        entry.WriteToFile(destFilePath, new ExtractionOptions
                        {
                            ExtractFullPath = false, // 关键：忽略压缩包内的目录层级
                            Overwrite = true
                        });

                        fileCount++;
                        AddLog($"解压文件：{Path.GetFileName(entry.Key)} → {archiveKey}");
                    }
                    catch (Exception ex)
                    {
                        AddLog($"解压文件 {Path.GetFileName(entry.Key)} 失败：{ex.Message}", LogLevel.Error);
                        continue; // 单个文件失败不中断整体解压
                    }
                }
            }
            return fileCount;
        }

        /// <summary>
        /// 处理文件夹（直接遍历文件并归档）
        /// </summary>
        private int ProcessFolder(string folderPath)
        {
            int fileCount = 0;
            // 递归遍历文件夹下的所有文件（含子目录）
            var allFiles = Directory.GetFiles(folderPath, "*.*", SearchOption.AllDirectories);

            foreach (var file in allFiles)
            {
                try
                {
                    // 跳过空文件
                    if (new FileInfo(file).Length == 0)
                        continue;

                    // 按选中规则生成归档Key（使用文件完整路径/文件名）
                    var archiveKey = SelectedFilingRule.GetArchiveKey(file);
                    var archiveFolder = Path.Combine(_tempArchiveDir, archiveKey);
                    // 创建归档文件夹
                    if (!Directory.Exists(archiveFolder))
                    {
                        Directory.CreateDirectory(archiveFolder);
                        AddLog($"创建归档文件夹：{archiveKey}");
                    }

                    // 复制文件到归档文件夹
                    var destFilePath = Path.Combine(archiveFolder, Path.GetFileName(file));
                    File.Copy(file, destFilePath, true);

                    fileCount++;
                    AddLog($"复制文件：{Path.GetFileName(file)} → {archiveKey}");
                }
                catch (Exception ex)
                {
                    AddLog($"复制文件 {Path.GetFileName(file)} 失败：{ex.Message}", LogLevel.Error);
                    continue; // 单个文件失败不中断整体处理
                }
            }
            return fileCount;
        }

        /// <summary>
        /// Tesseract OCR识别图片并生成Excel（支持模板/动态生成）
        /// </summary>
        private void OcrAndGenerateExcel()
        {
            try
            {
                if (!Directory.Exists(_tempArchiveDir) || Directory.GetDirectories(_tempArchiveDir).Length == 0)
                {
                    MessageBox.Show("请先完成源路径处理归档！", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                AddLog("开始OCR识别图片并生成Excel...");
                // 初始化Tesseract引擎（需确保TessData文件夹存在）
                var tessDataPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TessData");
                if (!Directory.Exists(tessDataPath))
                {
                    throw new Exception("未找到TessData语言包目录，请确认已放置chi_sim.traineddata/eng.traineddata");
                }

                // 处理中英双语选项
                var langCode = SelectedOcrLanguage switch
                {
                    "简体中文" => "chi_sim",
                    "英文" => "eng",
                    "中英双语" => "chi_sim+eng",
                    _ => "chi_sim"
                };

                IWorkbook workbook = null;
                ISheet sheet = null;
                int rowIndex = 1;

                // 判断是否使用Excel模板
                if (!string.IsNullOrEmpty(ExcelTemplateFilePath) && File.Exists(ExcelTemplateFilePath))
                {
                    // 有模板：读取模板文件
                    AddLog($"使用Excel模板：{Path.GetFileName(ExcelTemplateFilePath)}");
                    using (var fs = new FileStream(ExcelTemplateFilePath, FileMode.Open, FileAccess.Read))
                    {
                        workbook = ExcelTemplateFilePath.EndsWith(".xlsx")
                            ? new XSSFWorkbook(fs)
                            : (IWorkbook)new HSSFWorkbook(fs);
                    }
                    // 取第一个工作表
                    sheet = workbook.GetSheetAt(0);
                    // 模板表头行索引默认0，从第1行开始填充
                    rowIndex = 1;
                }
                else
                {
                    // 无模板：动态创建Excel
                    AddLog("未检测到Excel模板，动态生成Excel结构");
                    workbook = new XSSFWorkbook();
                    sheet = workbook.CreateSheet("OCR识别结果");
                }

                // 遍历所有归档文件夹
                foreach (var archiveFolder in Directory.GetDirectories(_tempArchiveDir))
                {
                    var folderName = Path.GetFileName(archiveFolder);
                    // 遍历图片文件（仅处理常见图片格式）
                    var imageFiles = Directory.GetFiles(archiveFolder)
                        .Where(f => new[] { ".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".gif" }.Contains(Path.GetExtension(f).ToLower()));

                    foreach (var imageFile in imageFiles)
                    {
                        var fileName = Path.GetFileName(imageFile);
                        AddLog($"识别图片：{fileName}");

                        // Tesseract OCR识别（单个图片异常不中断）
                        string ocrText = string.Empty;
                        string ocrStatus = "成功";
                        try
                        {
                            using (var engine = new TesseractEngine(tessDataPath, langCode, EngineMode.Default))
                            {
                                // 优化OCR识别精度
                                engine.SetVariable("tessedit_char_whitelist", "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ\u4e00-\u9fa5");
                                using (var pix = Pix.LoadFromFile(imageFile))
                                using (var page = engine.Process(pix))
                                {
                                    ocrText = page.GetText().Trim();
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            ocrText = $"识别失败：{ex.Message}";
                            ocrStatus = "失败";
                            AddLog($"识别图片 {fileName} 失败：{ex.Message}", LogLevel.Error);
                        }

                        // 解析OCR文本为键值对
                        var ocrKeyValue = ParseOcrTextToKeyValue(ocrText);
                        if (ocrKeyValue.Count == 0)
                        {
                            AddLog($"[{fileName}] 未识别到有效键值对，仅记录原始文本", LogLevel.Warn);
                            ocrKeyValue.Add("原始文本", ocrText);
                            ocrKeyValue.Add("识别状态", ocrStatus);
                            ocrKeyValue.Add("识别时间", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                        }

                        // 根据有无模板选择填充方式
                        if (!string.IsNullOrEmpty(ExcelTemplateFilePath) && File.Exists(ExcelTemplateFilePath))
                        {
                            // 有模板：按模板列名匹配填充
                            FillExcelWithTemplate(sheet, ref rowIndex, folderName, fileName, ocrKeyValue);
                        }
                        else
                        {
                            // 无模板：动态生成表头+填充内容
                            FillExcelDynamically(sheet, ref rowIndex, folderName, fileName, ocrKeyValue);
                        }

                        AddLog($"识别完成：{fileName} → 提取键值对数量：{ocrKeyValue.Count}，状态：{ocrStatus}");
                    }
                }

                // 自动调整列宽
                // 根据首行（表头）获取列数
                IRow headerRow = sheet.GetRow(0);
                if (headerRow != null)
                {
                    int lastCellNum = headerRow.LastCellNum;
                    for (int i = 0; i < lastCellNum; i++)
                    {
                        sheet.AutoSizeColumn(i);
                        // 限制最大列宽，避免过宽
                        var currentWidth = sheet.GetColumnWidth(i);
                        sheet.SetColumnWidth(i, currentWidth > 6000 ? 6000 : currentWidth);
                    }
                }

                // 保存Excel到临时目录
                _excelFilePath = Path.Combine(_tempArchiveDir, $"OCR识别结果_{DateTime.Now:yyyyMMddHHmmss}.xlsx");
                using (var fs = new FileStream(_excelFilePath, FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(fs);
                }

                AddLog($"Excel生成完成：{Path.GetFileName(_excelFilePath)}");
            }
            catch (Exception ex)
            {
                AddLog($"OCR识别/Excel生成失败：{ex.Message}", LogLevel.Error);
                MessageBox.Show($"识别失败：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 选择导出文件夹
        /// </summary>
        private void SelectExportFolder()
        {
            // 兼容WPF的文件夹选择对话框
            var folderDialog = new OpenFolderDialog
            {
                Title = "选择归档文件导出位置",
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            };

            if (folderDialog.ShowDialog() == true)
            {
                ExportFolderPath = folderDialog.FolderName;
                OnPropertyChanged(nameof(ExportFolderPath));
                AddLog($"已选择导出位置：{ExportFolderPath}");
            }
        }

        /// <summary>
        /// 导出归档文件夹和Excel文件
        /// </summary>
        private void ExportFiles()
        {
            try
            {
                if (string.IsNullOrEmpty(ExportFolderPath) || !Directory.Exists(ExportFolderPath))
                {
                    MessageBox.Show("请先选择有效的导出位置！", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                AddLog("开始导出文件...");
                // 生成最终导出文件夹（带时间戳避免覆盖）
                var exportDir = Path.Combine(ExportFolderPath, $"文档归档结果_{DateTime.Now:yyyyMMddHHmmss}");
                Directory.CreateDirectory(exportDir);

                // 复制归档文件夹
                foreach (var folder in Directory.GetDirectories(_tempArchiveDir))
                {
                    var destFolder = Path.Combine(exportDir, Path.GetFileName(folder));
                    CopyDirectory(folder, destFolder);
                    AddLog($"复制归档文件夹：{Path.GetFileName(folder)} → {exportDir}");
                }

                // 复制Excel文件（如果存在）
                if (!string.IsNullOrEmpty(_excelFilePath) && File.Exists(_excelFilePath))
                {
                    var excelDestPath = Path.Combine(exportDir, Path.GetFileName(_excelFilePath));
                    File.Copy(_excelFilePath, excelDestPath, true);
                    AddLog($"复制Excel文件：{Path.GetFileName(_excelFilePath)}");
                }

                AddLog($"导出完成！导出目录：{exportDir}");
                MessageBox.Show($"导出成功！\n目录：{exportDir}", "成功", MessageBoxButton.OK, MessageBoxImage.Information);
                // 自动打开导出目录
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = exportDir,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                AddLog($"导出失败：{ex.Message}", LogLevel.Error);
                MessageBox.Show($"导出失败：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        #endregion

        #region 辅助方法（Excel模板/动态生成核心逻辑）
        /// <summary>
        /// 解析OCR文本为键值对（支持"键：值""键=值""键-值"等格式）
        /// </summary>
        private Dictionary<string, string> ParseOcrTextToKeyValue(string ocrText)
        {
            var keyValueDict = new Dictionary<string, string>();
            if (string.IsNullOrEmpty(ocrText)) return keyValueDict;

            // 正则匹配：键（中文/英文/数字） + 分隔符（：/:/=/—/-） + 值（任意字符）
            var regex = new Regex(@"([\u4e00-\u9fa5a-zA-Z0-9]+)\s*[:=—-]\s*([^\n\r]+)", RegexOptions.Multiline);
            var matches = regex.Matches(ocrText);

            foreach (Match match in matches)
            {
                if (match.Groups.Count >= 3)
                {
                    var key = match.Groups[1].Value.Trim();
                    var value = match.Groups[2].Value.Trim();
                    if (!keyValueDict.ContainsKey(key))
                    {
                        keyValueDict.Add(key, value);
                    }
                }
            }

            return keyValueDict;
        }

        /// <summary>
        /// 有模板时填充Excel（按列名匹配）
        /// </summary>
        private void FillExcelWithTemplate(ISheet sheet, ref int rowIndex, string folderName, string fileName, Dictionary<string, string> ocrKeyValue)
        {
            // 获取模板表头（第0行）
            var headerRow = sheet.GetRow(0);
            if (headerRow == null)
            {
                throw new Exception("Excel模板无表头行，请检查模板格式");
            }

            // 创建新行（若已存在则复用，否则新建）
            var row = sheet.GetRow(rowIndex) ?? sheet.CreateRow(rowIndex);
            rowIndex++;

            // 遍历模板列，匹配键值对填充
            for (int colIndex = 0; colIndex < headerRow.LastCellNum; colIndex++)
            {
                var headerCell = headerRow.GetCell(colIndex);
                if (headerCell == null) continue;

                var headerText = headerCell.ToString()?.Trim();
                if (string.IsNullOrEmpty(headerText)) continue;

                // 特殊列：归档文件夹/图片文件名（固定匹配）
                if (headerText.Equals("归档文件夹", StringComparison.OrdinalIgnoreCase))
                {
                    row.CreateCell(colIndex).SetCellValue(folderName);
                }
                else if (headerText.Equals("图片文件名", StringComparison.OrdinalIgnoreCase))
                {
                    row.CreateCell(colIndex).SetCellValue(fileName);
                }
                // 匹配OCR识别的键值对
                else if (ocrKeyValue.ContainsKey(headerText))
                {
                    row.CreateCell(colIndex).SetCellValue(ocrKeyValue[headerText]);
                }
                // 无匹配值则留空
                else
                {
                    row.CreateCell(colIndex).SetCellValue("");
                }
            }
        }

        /// <summary>
        /// 无模板时动态生成Excel（自动创建表头+填充）
        /// </summary>
        private void FillExcelDynamically(ISheet sheet, ref int rowIndex, string folderName, string fileName, Dictionary<string, string> ocrKeyValue)
        {
            // 首次生成时创建表头
            if (sheet.GetRow(0) == null)
            {
                var headerRow = sheet.CreateRow(0);
                int colIndex = 0;

                // 固定表头列
                headerRow.CreateCell(colIndex++).SetCellValue("归档文件夹");
                headerRow.CreateCell(colIndex++).SetCellValue("图片文件名");

                // 动态添加OCR识别的键作为表头
                foreach (var key in ocrKeyValue.Keys)
                {
                    headerRow.CreateCell(colIndex++).SetCellValue(key);
                }
            }

            // 创建数据行
            var row = sheet.CreateRow(rowIndex);
            rowIndex++;
            int dataColIndex = 0;

            // 填充固定列
            row.CreateCell(dataColIndex++).SetCellValue(folderName);
            row.CreateCell(dataColIndex++).SetCellValue(fileName);

            // 填充OCR识别的键值对
            foreach (var value in ocrKeyValue.Values)
            {
                row.CreateCell(dataColIndex++).SetCellValue(value);
            }
        }
        #endregion

        #region 原有辅助方法
        /// <summary>
        /// 递归复制文件夹（含子目录）
        /// </summary>
        private void CopyDirectory(string sourceDir, string destDir)
        {
            Directory.CreateDirectory(destDir);
            // 复制文件
            foreach (var file in Directory.GetFiles(sourceDir))
            {
                try
                {
                    File.Copy(file, Path.Combine(destDir, Path.GetFileName(file)), true);
                }
                catch (Exception ex)
                {
                    AddLog($"复制文件 {Path.GetFileName(file)} 失败：{ex.Message}", LogLevel.Error);
                }
            }
            // 递归复制子目录
            foreach (var subDir in Directory.GetDirectories(sourceDir))
            {
                CopyDirectory(subDir, Path.Combine(destDir, Path.GetFileName(subDir)));
            }
        }

        /// <summary>
        /// 添加日志（适配项目现有LogHelper）
        /// </summary>
        public void AddLog(string content, LogLevel level = LogLevel.Info)
        {
            Application.Current.Dispatcher.Invoke(() =>
            {
                string displayContent = $"[{DateTime.Now:HH:mm:ss}] [{level}] {content}";
                LogList.Add(new UILogItem { Content = displayContent, Level = level });
            });
        }

        /// <summary>
        /// 清理目录（删除所有文件和子目录）
        /// </summary>
        private void CleanDirectory(string dirPath)
        {
            if (Directory.Exists(dirPath))
            {
                foreach (var file in Directory.GetFiles(dirPath))
                {
                    File.Delete(file);
                }
                foreach (var subDir in Directory.GetDirectories(dirPath))
                {
                    Directory.Delete(subDir, true);
                }
            }
        }

        /// <summary>
        /// 递归清理空文件夹
        /// </summary>
        private void CleanEmptyFolders(string dirPath)
        {
            if (!Directory.Exists(dirPath)) return;

            foreach (var subDir in Directory.GetDirectories(dirPath))
            {
                CleanEmptyFolders(subDir);
                // 如果子目录为空则删除
                if (Directory.GetFiles(subDir).Length == 0 && Directory.GetDirectories(subDir).Length == 0)
                {
                    Directory.Delete(subDir);
                    AddLog($"清理空文件夹：{subDir}");
                }
            }
        }

        /// <summary>
        /// 释放资源（清理临时目录）
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (_disposed) return;

            if (disposing)
            {
                // 清理托管资源
                if (Directory.Exists(_tempArchiveDir))
                {
                    try
                    {
                        Directory.Delete(_tempArchiveDir, true);
                        AddLog($"清理临时目录：{_tempArchiveDir}");
                    }
                    catch (Exception ex)
                    {
                        AddLog($"清理临时目录失败：{ex.Message}", LogLevel.Error);
                    }
                }
            }

            _disposed = true;
        }

        ~DocumentFilingViewModel()
        {
            Dispose(false);
        }
        #endregion
    }
}