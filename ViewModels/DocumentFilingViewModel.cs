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

        private bool _isAlreadyArchived;
        public bool IsAlreadyArchived
        {
            get => _isAlreadyArchived;
            set { _isAlreadyArchived = value; OnPropertyChanged(); }
        }

        // 临时归档目录（避免磁盘残留）
        private readonly string _tempArchiveDir;
        // 生成的Excel路径
        private string _excelFilePath;
        // 标记是否已释放资源
        private bool _disposed = false;
        #endregion

        #region 命令定义
        public ICommand SelectSourcePathCommand { get; }
        public ICommand ProcessSourceAndArchiveCommand { get; }
        public ICommand OcrAndGenerateExcelCommand { get; }
        public ICommand SelectExportFolderCommand { get; }
        public ICommand ExportFilesCommand { get; }
        public ICommand SelectExcelTemplateCommand { get; }
        public ICommand ClearLogCommand { get; }
        #endregion

        public DocumentFilingViewModel()
        {
            // 初始化日志集合
            LogList = new ObservableCollection<UILogItem>();
            // 初始化临时目录（带GUID避免冲突）
            _tempArchiveDir = Path.Combine(Path.GetTempPath(), $"DocArchive_{Guid.NewGuid():N}");
            Directory.CreateDirectory(_tempArchiveDir);
            // 默认导出路径为桌面
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            ExportFolderPath = desktopPath;

            // 初始化命令
            SelectSourcePathCommand = new RelayCommand(SelectSourcePath);
            ProcessSourceAndArchiveCommand = new RelayCommand(ProcessSourceAndArchive);
            OcrAndGenerateExcelCommand = new RelayCommand(OcrAndGenerateExcel);
            SelectExportFolderCommand = new RelayCommand(SelectExportFolder);
            ExportFilesCommand = new RelayCommand(ExportFiles);
            SelectExcelTemplateCommand = new RelayCommand(SelectExcelTemplate);
            ClearLogCommand = new RelayCommand(ExecuteClearLog);

            // 初始化归档规则（优先级从高到低）
            InitFilingRules();
            // 初始化OCR语言选项
            InitOcrLanguages();

            AddLog("文档归档工具初始化完成");
        }

        #region 初始化方法
        /// <summary>
        /// 初始化归档规则（优先级：下划线首段 > 姓名+身份证号 > 唯一字符串 > 文件名兜底）
        /// </summary>
        private void InitFilingRules()
        {
            FilingRuleList = new ObservableCollection<FilingRule>
            {
                // 规则0（最高优先级）：下划线分隔取首段（适配示例文件名：20200915235025178252491031_xxx）
                new FilingRule
                {
                    RuleName = "0. 下划线分隔首段（最高优先）",
                    GetArchiveKey = fileName =>
                    {
                        // 先剥离扩展名，避免扩展名干扰
                        string nameWithoutExt = Path.GetFileNameWithoutExtension(fileName);
                        if (string.IsNullOrWhiteSpace(nameWithoutExt))
                            return GetUniqueStringKey(fileName); // 降级到下一级规则

                        // 按下划线拆分，取第一个非空分段
                        if (nameWithoutExt.Contains("_"))
                        {
                            string core = nameWithoutExt.Split('_')[0].Trim();
                            if (!string.IsNullOrWhiteSpace(core))
                                return core;
                        }
                        // 匹配不到则降级到身份证规则
                        return GetIdCardKey(fileName);
                    }
                },
                // 规则1：匹配姓名+18位身份证号
                new FilingRule
                {
                    RuleName = "1. 姓名+18位身份证号",
                    GetArchiveKey = fileName =>
                    {
                        string idCardKey = GetIdCardKey(fileName);
                        if (!string.IsNullOrWhiteSpace(idCardKey) && idCardKey != GetUniqueStringKey(fileName))
                            return idCardKey;
                        // 降级到唯一字符串规则
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
        /// 提取身份证号Key
        /// </summary>
        private string GetIdCardKey(string fileName)
        {
            string nameWithoutExt = Path.GetFileNameWithoutExtension(fileName);
            if (string.IsNullOrWhiteSpace(nameWithoutExt))
                return string.Empty;

            // 严格匹配18位身份证号（最后一位支持X/x）
            var idCardRegex = new Regex(@"[1-9]\d{5}(19|20)\d{2}((0[1-9])|(1[0-2]))(([0-2][1-9])|10|20|30|31)\d{3}[\dXx]", RegexOptions.Compiled);
            var idCardMatch = idCardRegex.Match(nameWithoutExt);
            if (idCardMatch.Success)
            {
                // 提取身份证号前的姓名（仅中文，最多4个汉字）
                var nameSegment = nameWithoutExt.Substring(0, idCardMatch.Index);
                var nameMatch = Regex.Match(nameSegment, @"[\u4e00-\u9fa5]{1,4}");
                var name = nameMatch.Success ? nameMatch.Value.Trim() : string.Empty;
                return string.IsNullOrEmpty(name) ? idCardMatch.Value : $"{name}_{idCardMatch.Value}";
            }
            return string.Empty;
        }

        /// <summary>
        /// 提取唯一字符串Key（支持多分隔符）
        /// </summary>
        private string GetUniqueStringKey(string fileName)
        {
            string nameWithoutExt = Path.GetFileNameWithoutExtension(fileName);
            if (string.IsNullOrWhiteSpace(nameWithoutExt))
                return Path.GetFileName(fileName) ?? string.Empty;

            // 先处理下划线（兜底）
            if (nameWithoutExt.Contains("_"))
            {
                string core = nameWithoutExt.Split('_')[0].Trim();
                if (!string.IsNullOrWhiteSpace(core))
                    return core;
            }

            // 匹配连续的字母/数字/下划线（长度≥6）
            var uniqueStrMatch = Regex.Match(nameWithoutExt, @"(?<=[^a-zA-Z0-9_])[a-zA-Z0-9_]{6,}(?=[^a-zA-Z0-9_])");
            if (uniqueStrMatch.Success)
                return uniqueStrMatch.Value;

            // 补充：按常见分隔符（-、空格、#）拆分取首段
            char[] separators = new[] { '-', ' ', '#' };
            foreach (char sep in separators)
            {
                if (nameWithoutExt.Contains(sep))
                {
                    string core = nameWithoutExt.Split(sep)[0].Trim();
                    if (!string.IsNullOrWhiteSpace(core))
                        return core;
                }
            }

            // 最终兜底：返回文件名主体
            return nameWithoutExt;
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
        /// 选择源路径（勾选已归档时仅允许选文件夹）
        /// </summary>
        private void SelectSourcePath()
        {
            // 勾选已归档模式时，强制只能选文件夹
            if (IsAlreadyArchived)
            {
                var folderDialog = new OpenFolderDialog
                {
                    Title = "选择已归档的文件夹",
                    InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
                };

                if (folderDialog.ShowDialog() == true)
                {
                    SourcePath = folderDialog.FolderName;
                    AddLog($"已选择已归档文件夹：{SourcePath}");
                }
                return;
            }

            // 未勾选时，保留原逻辑（选压缩包/文件夹）
            var result = MessageBox.Show("请选择要处理的类型：\n【是】- 压缩包文件（zip/rar/7z）\n【否】- 文件夹", "选择处理类型",
                MessageBoxButton.YesNoCancel, MessageBoxImage.Question);

            if (result == MessageBoxResult.Cancel)
                return;

            if (result == MessageBoxResult.Yes)
            {
                // 选择压缩包文件逻辑
                var openFileDialog = new OpenFileDialog
                {
                    Title = "选择压缩包文件",
                    Filter = "压缩包文件|*.zip;*.rar;*.7z|所有文件|*.*",
                    RestoreDirectory = true
                };

                if (openFileDialog.ShowDialog() == true)
                {
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
                // 选择文件夹逻辑
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
                // 勾选已归档时跳过归档
                if (IsAlreadyArchived)
                {
                    // 校验源路径是否为文件夹
                    if (!Directory.Exists(SourcePath))
                    {
                        MessageBox.Show("已归档模式下，源路径必须是文件夹！", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }
                    AddLog("已勾选「已归档文件夹」，跳过归档步骤，可直接执行OCR识别");
                    MessageBox.Show("已归档模式：跳过归档步骤，请直接点击「OCR识别并生成Excel」按钮", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

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

                // 区分源路径类型（文件/文件夹）
                if (File.Exists(SourcePath))
                {
                    // 源路径是压缩包文件 → 解压并归档
                    fileCount = ProcessCompressFile(SourcePath);
                    if (fileCount == 0)
                    {
                        AddLog("压缩包处理过程中检测到错误，已终止！", LogLevel.Error);
                        return;
                    }
                }
                else if (Directory.Exists(SourcePath))
                {
                    // 源路径是文件夹 → 直接遍历并归档
                    fileCount = ProcessFolder(SourcePath);
                    if (fileCount == 0)
                    {
                        AddLog("文件夹处理过程中检测到错误，已终止！", LogLevel.Error);
                        return;
                    }
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
        /// 处理压缩包文件（解压并归档，触发错误立即终止）
        /// </summary>
        private int ProcessCompressFile(string compressFilePath)
        {
            int fileCount = 0;

            using (var archive = ArchiveFactory.Open(compressFilePath))
            {
                foreach (var entry in archive.Entries)
                {
                    if (entry.IsDirectory || entry.Size == 0) continue;

                    try
                    {
                        string fileName = Path.GetFileName(entry.Key);
                        string archiveKey;

                        if (IsAlreadyArchived)
                        {
                            archiveKey = Path.GetFileNameWithoutExtension(fileName) ?? string.Empty;
                            AddLog($"已归档模式：文件「{fileName}」直接使用文件名作为Key → {archiveKey}");
                        }
                        else
                        {
                            archiveKey = SelectedFilingRule.GetArchiveKey(fileName);
                            // 空Key校验：触发错误则弹窗并终止方法
                            if (string.IsNullOrWhiteSpace(archiveKey))
                            {
                                string errorMsg = $"文件无法通过当前规则「{SelectedFilingRule.RuleName}」提取归档Key，请更换规则！";
                                AddLog(errorMsg, LogLevel.Error);

                                // 弹窗后直接return，终止所有后续处理
                                Application.Current.Dispatcher.Invoke(() =>
                                    MessageBox.Show(errorMsg, "提示", MessageBoxButton.OK, MessageBoxImage.Warning));
                                return fileCount;
                            }
                        }

                        var archiveFolder = Path.Combine(_tempArchiveDir, archiveKey);
                        if (!Directory.Exists(archiveFolder))
                        {
                            Directory.CreateDirectory(archiveFolder);
                            AddLog($"创建归档文件夹：{archiveKey}");
                        }

                        var destFilePath = GetUniqueDestFilePath(archiveFolder, fileName);
                        entry.WriteToFile(destFilePath, new ExtractionOptions
                        {
                            ExtractFullPath = false,
                            Overwrite = true
                        });

                        fileCount++;
                        AddLog($"解压文件：{fileName} → {archiveKey}");
                    }
                    catch (Exception ex)
                    {
                        string errorMsg = $"解压文件 {Path.GetFileName(entry.Key)} 失败：{ex.Message}";
                        AddLog(errorMsg, LogLevel.Error);

                        // 异常时弹窗并终止
                        Application.Current.Dispatcher.Invoke(() =>
                            MessageBox.Show(errorMsg, "解压失败", MessageBoxButton.OK, MessageBoxImage.Error));
                        return 0;
                    }
                }
            }
            return fileCount;
        }

        /// <summary>
        /// 处理文件夹（直接遍历文件并归档，触发错误立即终止）
        /// </summary>
        private int ProcessFolder(string folderPath)
        {
            int fileCount = 0;
            var files = Directory.GetFiles(folderPath, "*.*", SearchOption.AllDirectories);

            foreach (var file in files)
            {
                if (new FileInfo(file).Length == 0) continue;

                string fileName = Path.GetFileName(file);
                string archiveKey;

                try
                {
                    if (IsAlreadyArchived)
                    {
                        archiveKey = Path.GetFileNameWithoutExtension(fileName) ?? string.Empty;
                        AddLog($"已归档模式：文件「{fileName}」直接使用文件名作为Key → {archiveKey}");
                    }
                    else
                    {
                        archiveKey = SelectedFilingRule.GetArchiveKey(fileName);
                        if (string.IsNullOrWhiteSpace(archiveKey))
                        {
                            string errorMsg = $"文件无法通过当前规则「{SelectedFilingRule.RuleName}」提取归档Key，请更换规则！";
                            AddLog(errorMsg, LogLevel.Error);

                            Application.Current.Dispatcher.Invoke(() =>
                                MessageBox.Show(errorMsg, "提示", MessageBoxButton.OK, MessageBoxImage.Warning));
                            return fileCount;
                        }
                    }

                    var archiveFolder = Path.Combine(_tempArchiveDir, archiveKey);
                    if (!Directory.Exists(archiveFolder))
                    {
                        Directory.CreateDirectory(archiveFolder);
                        AddLog($"创建归档文件夹：{archiveKey}");
                    }

                    var destFilePath = GetUniqueDestFilePath(archiveFolder, fileName);
                    File.Copy(file, destFilePath, true);

                    fileCount++;
                    AddLog($"复制文件：{fileName} → {archiveKey}");
                }
                catch (Exception ex)
                {
                    string errorMsg = $"复制文件 {fileName} 失败：{ex.Message}";
                    AddLog(errorMsg, LogLevel.Error);

                    // 异常时弹窗并终止
                    Application.Current.Dispatcher.Invoke(() =>
                        MessageBox.Show(errorMsg, "复制失败", MessageBoxButton.OK, MessageBoxImage.Error));
                    return 0;
                }
            }
            return fileCount;
        }

        /// <summary>
        /// Tesseract OCR识别图片并生成Excel（支持模板/动态生成 + 兼容已归档文件夹）
        /// </summary>
        private void OcrAndGenerateExcel()
        {
            try
            {
                // ========== 核心修改：动态切换OCR读取目录 ==========
                string ocrRootDir = string.Empty;
                if (IsAlreadyArchived)
                {
                    // 已归档模式：读取SourcePath（必须是文件夹）
                    if (string.IsNullOrEmpty(SourcePath) || !Directory.Exists(SourcePath))
                    {
                        MessageBox.Show("已归档模式下，请先选择有效的文件夹！", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }
                    ocrRootDir = SourcePath;
                    AddLog($"已归档模式：直接读取文件夹 {ocrRootDir} 进行OCR识别");
                }
                else
                {
                    // 未归档模式：读取临时目录（需先执行归档）
                    if (!Directory.Exists(_tempArchiveDir) || Directory.GetDirectories(_tempArchiveDir).Length == 0)
                    {
                        MessageBox.Show("未归档模式下，请先点击「开始处理并归档」按钮完成归档！", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }
                    ocrRootDir = _tempArchiveDir;
                    AddLog($"未归档模式：读取临时归档目录 {ocrRootDir} 进行OCR识别");
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

                // ========== 修改：遍历动态切换的根目录 ==========
                foreach (var archiveFolder in Directory.GetDirectories(ocrRootDir))
                {
                    var folderName = Path.GetFileName(archiveFolder);
                    // 定义需要遍历的图片扩展名（可按需扩展）
                    var imageExtensions = new[] { ".jpg", ".jpeg", ".png", ".bmp", ".tiff", ".gif", ".webp" };
                    // 调用通用递归方法，获取所有嵌套层级的图片文件
                    var imageFiles = GetAllFilesRecursively(archiveFolder, imageExtensions);

                    // 日志打印：当前归档文件夹下找到的图片总数（含嵌套）
                    AddLog($"归档文件夹 {folderName} 下共找到 {imageFiles.Count} 张图片（含多层嵌套子文件夹）");

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
                IRow headerRow = sheet.GetRow(0);
                if (headerRow != null)
                {
                    int lastCellNum = headerRow.LastCellNum;
                    for (int i = 0; i < lastCellNum; i++)
                    {
                        sheet.AutoSizeColumn(i);
                        var currentWidth = sheet.GetColumnWidth(i);
                        sheet.SetColumnWidth(i, currentWidth > 6000 ? 6000 : currentWidth);
                    }
                }

                // Excel保存路径（兼容已归档模式）
                string excelSaveDir = string.IsNullOrEmpty(ExportFolderPath) ? _tempArchiveDir : ExportFolderPath;
                _excelFilePath = Path.Combine(excelSaveDir, $"OCR识别结果_{DateTime.Now:yyyyMMddHHmmss}.xlsx");
                using (var fs = new FileStream(_excelFilePath, FileMode.Create, FileAccess.Write))
                {
                    workbook.Write(fs);
                }

                AddLog($"Excel生成完成：{Path.GetFileName(_excelFilePath)}");
                MessageBox.Show($"Excel生成成功！\n路径：{_excelFilePath}", "成功", MessageBoxButton.OK, MessageBoxImage.Information);
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
        /// 导出归档文件夹和Excel文件（兼容已归档模式：仅导出Excel）
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

                // 校验Excel是否生成
                if (string.IsNullOrEmpty(_excelFilePath) || !File.Exists(_excelFilePath))
                {
                    MessageBox.Show("未检测到生成的Excel文件，请先执行OCR识别！", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                AddLog("开始导出文件...");
                // 生成最终导出文件夹（带时间戳避免覆盖）
                var exportDir = Path.Combine(ExportFolderPath, $"文档归档结果_{DateTime.Now:yyyyMMddHHmmss}");
                Directory.CreateDirectory(exportDir);

                // IsAlreadyArchived切换导出逻辑 
                if (IsAlreadyArchived)
                {
                    // 已归档模式：仅复制Excel文件
                    var excelDestPath = Path.Combine(exportDir, Path.GetFileName(_excelFilePath));
                    File.Copy(_excelFilePath, excelDestPath, true);
                    AddLog($"仅导出Excel文件：{Path.GetFileName(_excelFilePath)} → {exportDir}");
                }
                else
                {
                    // 未归档模式：复制归档文件夹 + Excel（原有逻辑）
                    // 复制归档文件夹
                    foreach (var folder in Directory.GetDirectories(_tempArchiveDir))
                    {
                        var destFolder = Path.Combine(exportDir, Path.GetFileName(folder));
                        CopyDirectory(folder, destFolder);
                        AddLog($"复制归档文件夹：{Path.GetFileName(folder)} → {exportDir}");
                    }

                    // 复制Excel文件
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

        #region 辅助方法           
        /// <summary>
        /// 获取唯一的目标文件路径（处理重名，自动加数字后缀）
        /// </summary>
        private string GetUniqueDestFilePath(string targetFolder, string fileName)
        {
            string destFilePath = Path.Combine(targetFolder, fileName);
            if (!File.Exists(destFilePath))
                return destFilePath;

            // 重名则添加数字后缀（如：文件名_1.jpg）
            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(fileName);
            string ext = Path.GetExtension(fileName);
            int suffix = 1;
            while (File.Exists(destFilePath))
            {
                destFilePath = Path.Combine(targetFolder, $"{fileNameWithoutExt}_{suffix}{ext}");
                suffix++;
            }
            return destFilePath;
        }

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
        /// 通用型递归文件获取方法（支持多层嵌套子文件夹，适配任意文件类型）
        /// </summary>
        /// <param name="rootDir">根目录</param>
        /// <param name="fileExtensions">需要匹配的文件扩展名列表（如：new[] { ".jpg", ".pdf" }）</param>
        /// <param name="ignoreCase">是否忽略扩展名大小写（默认true）</param>
        /// <returns>所有匹配文件的完整路径列表</returns>
        private List<string> GetAllFilesRecursively(string rootDir, string[] fileExtensions, bool ignoreCase = true)
        {
            var matchedFiles = new List<string>();
            // 校验根目录有效性
            if (!Directory.Exists(rootDir))
            {
                AddLog($"根目录不存在：{rootDir}", LogLevel.Warn);
                return matchedFiles;
            }

            try
            {
                // 1. 获取当前目录下匹配的文件（兼容大小写）
                var currentFiles = Directory.GetFiles(rootDir)
                    .Where(file =>
                    {
                        string fileExt = Path.GetExtension(file);
                        // 空扩展名（如无后缀文件）直接过滤
                        if (string.IsNullOrEmpty(fileExt)) return false;
                        // 按配置决定是否忽略大小写
                        return ignoreCase
                            ? fileExtensions.Contains(fileExt.ToLower())
                            : fileExtensions.Contains(fileExt);
                    })
                    .ToList();
                matchedFiles.AddRange(currentFiles);

                // 2. 递归遍历所有子目录（无论嵌套多少层）
                foreach (var subDir in Directory.GetDirectories(rootDir))
                {
                    var subDirFiles = GetAllFilesRecursively(subDir, fileExtensions, ignoreCase);
                    matchedFiles.AddRange(subDirFiles);
                }
            }
            catch (UnauthorizedAccessException ex)
            {
                AddLog($"无权限访问目录 {rootDir}：{ex.Message}", LogLevel.Warn);
            }
            catch (PathTooLongException ex)
            {
                AddLog($"目录路径过长无法访问 {rootDir}：{ex.Message}", LogLevel.Warn);
            }
            catch (Exception ex)
            {
                AddLog($"遍历目录 {rootDir} 失败：{ex.Message}", LogLevel.Warn);
            }

            return matchedFiles;
        }

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

        /// <summary>
        /// 添加日志
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
        /// 清空日志
        /// </summary>
        private void ExecuteClearLog()
        {
            Application.Current.Dispatcher.Invoke(() =>
            {
                LogList.Clear();
            });
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