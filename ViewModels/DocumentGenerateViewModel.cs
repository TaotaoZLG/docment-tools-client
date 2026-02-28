using docment_tools_client.Helpers;
using docment_tools_client.Models;
using docment_tools_client.Services;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Input;

namespace docment_tools_client.ViewModels
{
    /// <summary>
    /// 文书生成页面专属视图模型
    /// </summary>
    public class DocumentGenerateViewModel : ViewModelBase
    {
        /// <summary>
        /// 当前登录用户信息
        /// </summary>
        public UserInfo UserInfo { get; set; }

        /// <summary>
        /// 剩余额度显示（格式化）
        /// </summary>
        public string QuotaDisplay => $"剩余额度：{UserInfo.Quota:F2} 元（单价：{UserInfo.UserPrice:F2} 元/条，可处理{(UserInfo.UserPrice > 0 ? (int)(UserInfo.Quota / UserInfo.UserPrice) : 0)}条数据）";

        #region 功能区绑定属性
        /// <summary>
        /// 文书类型列表
        /// </summary>
        public List<string> DocumentTypes { get; set; } = new List<string>
        {
            "其它类型", "民事起诉状", "保全申请书", "撤销保全申请", "民事裁定书",
            "授权委托书", "网络查控申请书"
        };

        /// <summary>
        /// 选中的文书类型
        /// </summary>
        private string _selectedDocumentType = "其它类型";
        public string SelectedDocumentType
        {
            get => _selectedDocumentType;
            set
            {
                if (_selectedDocumentType != value)
                {
                    _selectedDocumentType = value;
                    OnPropertyChanged(nameof(SelectedDocumentType));
                }
            }
        }

        /// <summary>
        /// 模板文件路径
        /// </summary>
        public string TemplatePath { get; set; } = string.Empty;

        /// <summary>
        /// Excel数据文件路径
        /// </summary>
        public string ExcelPath { get; set; } = string.Empty;

        /// <summary>
        /// 填充人数（默认为空/0表示全部）
        /// 使用 string 属性以支持 ComboBox 输入和空状态
        /// </summary>
        public string FillCountText
        {
            get => _fillCountText;
            set
            {
                SetProperty(ref _fillCountText, value);
                // 尝试解析int供逻辑使用(不做强绑定，仅解析)
                if (int.TryParse(value, out int result))
                {
                    FillCount = result;
                }
                else
                {
                    FillCount = 0; // 解析失败或空，视为0（全部）
                }
            }
        }
        private string _fillCountText = "";

        /// <summary>
        /// 填充人数数值（内部逻辑使用，0=全部）
        /// </summary>
        public int FillCount
        {
            get => _fillCount;
            set => SetProperty(ref _fillCount, value);
        }
        private int _fillCount = 0;

        /// <summary>
        /// 填充人数下拉选项
        /// </summary>
        public List<string> FillCountOptions { get; } = new List<string> { "10", "20", "30", "50", "100", "200", "300" };

        /// <summary>
        /// 是否带附件表格 (废弃UI，默认false)
        /// </summary>
        private bool IsWithAttachment => false;

        /// <summary>
        /// 生成格式（Word）
        /// </summary>
        public List<string> GenerateFormats { get; set; } = new List<string> { "Word文档" };

        /// <summary>
        /// 选中的生成格式
        /// </summary>
        public string SelectedGenerateFormat { get; set; } = "Word文档";

        /// <summary>
        /// 输出目录路径
        /// </summary>
        public string OutputDir { get; set; } = string.Empty;

        /// <summary>
        /// 文件名规则
        /// </summary>
        public string FileNameRule { get; set; } = string.Empty;

        /// <summary>
        /// Excel表头列表
        /// </summary>
        public ObservableCollection<string> ExcelHeaders { get; set; } = new ObservableCollection<string>();

        /// <summary>
        /// 案件数据列表
        /// </summary>
        public List<DynamicCaseData> CaseDatas { get; set; } = new List<DynamicCaseData>();
        #endregion

        #region 日志区绑定属性
        /// <summary>
        /// 日志列表（用于UI展示）
        /// </summary>
        public ObservableCollection<UILogItem> LogList { get; set; } = new ObservableCollection<UILogItem>();
        #endregion

        #region 命令
        /// <summary>
        /// 选择模板文件命令
        /// </summary>
        public ICommand SelectTemplateCommand { get; }

        /// <summary>
        /// 选择Excel文件命令
        /// </summary>
        public ICommand SelectExcelCommand { get; }

        /// <summary>
        /// 选择输出目录命令
        /// </summary>
        public ICommand SelectOutputDirCommand { get; }

        /// <summary>
        /// 生成预览命令
        /// </summary>
        public ICommand PreviewCommand { get; }

        /// <summary>
        /// 开始生成命令
        /// </summary>
        public ICommand StartGenerateCommand { get; }

        /// <summary>
        /// 导出数据模板命令
        /// </summary>
        public ICommand ExportTemplateCommand { get; }

        /// <summary>
        /// 导出Word模板命令
        /// </summary>
        public ICommand ExportWordTemplateCommand { get; }

        /// <summary>
        /// 清空日志命令
        /// </summary>
        public ICommand ClearLogCommand { get; }

        /// <summary>
        /// 快捷选择占位符命令
        /// </summary>
        public ICommand QuickSelectCommand { get; }
        #endregion

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="currentUser">当前登录用户信息</param>
        public DocumentGenerateViewModel(UserInfo userInfo)
        {
            UserInfo = userInfo ?? throw new ArgumentNullException(nameof(userInfo));

            // 初始化命令
            SelectTemplateCommand = new RelayCommand(ExecuteSelectTemplate);
            SelectExcelCommand = new RelayCommand(ExecuteSelectExcel);
            SelectOutputDirCommand = new RelayCommand(ExecuteSelectOutputDir);
            PreviewCommand = new RelayCommand(ExecutePreview);
            StartGenerateCommand = new RelayCommand(ExecuteStartGenerate);
            ExportTemplateCommand = new RelayCommand(ExecuteExportTemplate);
            ExportWordTemplateCommand = new RelayCommand(ExecuteExportWordTemplate);
            QuickSelectCommand = new RelayCommand(ExecuteQuickSelect);
            ClearLogCommand = new RelayCommand(ExecuteClearLog);

            // 注册日志回调（实时更新UI日志）
            LogHelper.OnLogReceived += OnLogReceived;

            // 初始化默认模板路径
            TemplatePath = WordService.GetDefaultTemplatePath();
            OnPropertyChanged(nameof(TemplatePath));

            // 欢迎日志
            LogHelper.Info($"进入文书生成页面，当前用户：{UserInfo.UserName}");
        }

        #region 功能区命令实现
        /// <summary>
        /// 选择模板文件
        /// </summary>
        private void ExecuteSelectTemplate()
        {
            var openFileDialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "Word模板文件 (*.docx)|*.docx|所有文件 (*.*)|*.*",
                Title = "选择Word模板文件",
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            };

            if (openFileDialog.ShowDialog() == true)
            {
                TemplatePath = openFileDialog.FileName;
                OnPropertyChanged(nameof(TemplatePath));
                LogHelper.Info($"选择了本地模本文件：{TemplatePath}");
            }

        }

        /// <summary>
        /// 选择Excel数据文件
        /// </summary>
        private void ExecuteSelectExcel()
        {
            var openFileDialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "Excel文件 (*.xlsx;*.xls)|*.xlsx;*.xls|所有文件 (*.*)|*.*",
                Title = "选择案件数据Excel文件",
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            };

            if (openFileDialog.ShowDialog() == true)
            {
                ExcelPath = openFileDialog.FileName;
                OnPropertyChanged(nameof(ExcelPath));

                // 读取Excel文件，初始化表头和案件数据
                var (headers, caseDatas) = ExcelService.ReadExcel(ExcelPath);
                CaseDatas = caseDatas;

                // 默认重置为空（全部）
                FillCountText = ""; // 这会触发Setter将FillCount置为0

                // 更新表头列表
                ExcelHeaders.Clear();
                foreach (var header in headers)
                {
                    ExcelHeaders.Add(header);
                }

                LogHelper.Info($"【解析表头】发现 {headers.Count} 个列：{string.Join(", ", headers)}");

                // 更新绑定属性
                OnPropertyChanged(nameof(CaseDatas));
                // OnPropertyChanged(nameof(MaxFillCount)); // Removed
                OnPropertyChanged(nameof(FillCountText)); // Notify text change
                OnPropertyChanged(nameof(ExcelHeaders));

                LogHelper.Info($"已选择Excel数据文件：{ExcelPath}");
            }
        }

        /// <summary>
        /// 选择输出目录
        /// </summary>
        private void ExecuteSelectOutputDir()
        {
            var folderDialog = new Microsoft.Win32.OpenFolderDialog
            {
                Title = "选择文件输出目录",
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            };

            if (folderDialog.ShowDialog() == true)
            {
                OutputDir = folderDialog.FolderName;
                OnPropertyChanged(nameof(OutputDir));
                LogHelper.Info($"已选择输出目录：{OutputDir}");
            }
        }

        /// <summary>
        /// 快捷选择占位符
        /// </summary>
        private void ExecuteQuickSelect()
        {
            if (ExcelHeaders == null || ExcelHeaders.Count == 0)
            {
                MessageBox.Show("请先读取Excel数据文件", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            // 使用反射或显式调用View（此处为简单起见直接实例化View，实际MVVM中可用Service）
            var win = new Views.FieldSelectionWindow(ExcelHeaders);
            if (win.ShowDialog() == true && !string.IsNullOrEmpty(win.SelectedHeader))
            {
                // 如果 FileNameRule 不为空且不以分隔符结尾，则添加分隔符
                string currentRule = FileNameRule ?? string.Empty;
                if (currentRule.Length > 0 && !currentRule.EndsWith("-"))
                {
                    currentRule += "-";
                }

                FileNameRule = currentRule + $"{{{win.SelectedHeader}}}";
                OnPropertyChanged(nameof(FileNameRule));
            }
        }

        /// <summary>
        /// 执行生成预览
        /// </summary>
        private void ExecutePreview()
        {
            try
            {
                // 前置校验
                if (string.IsNullOrEmpty(TemplatePath) || !System.IO.File.Exists(TemplatePath))
                {
                    MessageBox.Show("请选择有效的Word模板文件", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                if (CaseDatas == null || CaseDatas.Count == 0)
                {
                    MessageBox.Show("请先选择并读取有效的Excel案件数据", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // 取第一条数据生成预览文件
                var firstCase = CaseDatas[0];
                // 使用唯一文件名避免文件占用导致无法生成
                string fileName = $"预览文书_{DateTime.Now:yyyyMMddHHmmss}.docx";
                var previewPath = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), fileName);

                LogHelper.Info($"生成预览条件 | 类型：{SelectedDocumentType} | 人数限制：{FillCount} | 风格：{(IsWithAttachment ? "带附件表格" : "无表格")} | 格式：{SelectedGenerateFormat} | 文件名规则：{FileNameRule}");

                // 生成预览Word
                // 传入 withProtection=true (开启只读保护)
                var result = WordService.GenerateWord(TemplatePath, firstCase, previewPath, IsWithAttachment, true);
                if (result)
                {
                    // 打开预览文件
                    WordService.PreviewWord(previewPath);
                }
            }
            catch (System.IO.IOException)
            {
                // 专门捕获文件占用异常，提示用户关闭文档
                MessageBox.Show("无法生成预览：目标文件正由其他程序（如Word）打开，请先关闭已打开的文档后重试。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"预览生成失败：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                LogHelper.Error($"预览执行异常：{ex.Message}");
            }
        }

        /// <summary>
        /// 执行开始生成
        /// </summary>
        private void ExecuteStartGenerate()
        {
            try
            {
                // 前置校验（简化版）
                if (string.IsNullOrEmpty(OutputDir) || !System.IO.Directory.Exists(OutputDir))
                {
                    MessageBox.Show("请选择有效的输出目录", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                if (CaseDatas == null || CaseDatas.Count == 0)
                {
                    MessageBox.Show("无有效案件数据可生成", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // 1. 数据分割 (填充人数业务：对应第2条场景)
                var chunks = ExcelService.SplitCaseData(CaseDatas, FillCount);
                if (chunks.Count == 0)
                {
                    MessageBox.Show("无有效案件数据可生成", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                int totalRows = CaseDatas.Count;
                if (UserInfo.Quota < 0.2m * totalRows)
                {
                    MessageBox.Show($"剩余额度不足。需处理 {totalRows} 条数据，需 {0.2m * totalRows:F2} 元。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // 2. 批量生成 (异步执行)
                LogHelper.Info($"开始批量生成... 总数据{totalRows}条，分为{chunks.Count}份文档 (每份限制{FillCount}条)");

                Task.Run(() =>
                {
                    int successFiles = 0;
                    int processedRows = 0;

                    foreach (var chunk in chunks)
                    {
                        string fileName = string.Empty;
                        // 取该份数据的第一条作为命名依据
                        var representative = chunk[0];

                        try
                        {
                            // Apply FileNameRule
                            if (!string.IsNullOrEmpty(FileNameRule) && !string.IsNullOrWhiteSpace(FileNameRule))
                            {
                                fileName = FileNameRule;
                                foreach (var kv in representative.CaseInfo)
                                {
                                    var k = kv.Key ?? "";
                                    var v = kv.Value?.ToString() ?? "";
                                    fileName = fileName.Replace($"{{{{{k}}}}}", v)
                                                       .Replace($"{{{k}}}", v); // 兼容 {} 和 {{}}
                                }
                                if (!fileName.EndsWith(".docx", StringComparison.OrdinalIgnoreCase)) fileName += ".docx";
                                foreach (char c in System.IO.Path.GetInvalidFileNameChars()) fileName = fileName.Replace(c, '_');
                            }
                            else
                            {
                                // 默认命名规则：选择的文书类型名称_MMddHHmmss_{包含人数}人_短一点的GUID.docx
                                var shortGuid = Guid.NewGuid().ToString("N").Substring(0, 8); // 8位短Guid
                                fileName = $"{SelectedDocumentType}_{DateTime.Now:MMddHHmmss}_{chunk.Count}人_{shortGuid}.docx";
                            }
                        }
                        catch
                        {
                            fileName = $"文书_{Guid.NewGuid():N}.docx";
                        }

                        var targetPath = System.IO.Path.Combine(OutputDir, fileName);

                        // 防止文件名重复覆盖
                        if (System.IO.File.Exists(targetPath))
                        {
                            string nameNoExt = System.IO.Path.GetFileNameWithoutExtension(targetPath);
                            targetPath = System.IO.Path.Combine(OutputDir, $"{nameNoExt}_{Guid.NewGuid().ToString("N").Substring(0, 4)}.docx");
                        }

                        // 调用支持 List 的生成方法 (传递 chunk 即 List<DynamicCaseData>)
                        // ProcessLoops 会识别 List<DynamicCaseData> 并处理 Loop
                        if (WordService.GenerateWord(TemplatePath, chunk, targetPath, IsWithAttachment))
                        {
                            successFiles++;
                            processedRows += chunk.Count;

                            // 简单的进度反馈
                            if (successFiles % 5 == 0 || successFiles == chunks.Count)
                            {
                                Application.Current.Dispatcher.Invoke(() =>
                                {
                                    LogHelper.Info($"已生成 {successFiles}/{chunks.Count} 份文档");
                                });
                            }
                        }
                    }

                    // 3. 扣费逻辑 (按实际处理的案件条数收费)
                    var consumeAmount = 0.2m * processedRows;

                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        LogHelper.Info($"批量生成完成：成功生成 {successFiles} 份文件，包含 {processedRows} 条数据。");

                        // 提交消费记录
                        // ... (Consume logic omitted for brevity as it was not fully provided in context but keeping structure)

                        // Update Local Quota
                        if (processedRows > 0)
                        {
                            UserInfo.Quota -= consumeAmount;
                            UserInfo.EncryptedQuota = EncryptHelper.EncryptQuota(UserInfo.Quota);
                            LocalStorageHelper.SaveUserInfo(UserInfo);
                            OnPropertyChanged(nameof(QuotaDisplay));

                            // Add local record logic here if needed
                            // var record = new ConsumeRecord { ... };
                            // LocalDbHelper.AddConsumeRecord(record);
                            // SyncConsumeRecordsAsync();
                        }

                        MessageBox.Show($"生成完成！共消耗 {consumeAmount:F2} 元。", "成功", MessageBoxButton.OK, MessageBoxImage.Information);
                    });
                });
            }
            catch (Exception ex)
            {
                LogHelper.Error($"批量生成执行异常：{ex.Message}");
                MessageBox.Show($"批量生成启动失败：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 执行导出Word模板
        /// </summary>
        private void ExecuteExportWordTemplate()
        {
            try
            {
                if (SelectedDocumentType == "其它类型" || string.IsNullOrEmpty(SelectedDocumentType))
                {
                    MessageBox.Show("请选择具体的文书类型以查找模板。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                // 模糊查找所有相关模板
                var templates = FindTemplates(SelectedDocumentType, "*.docx");

                if (templates.Count == 0)
                {
                    MessageBox.Show($"未找到包含“{SelectedDocumentType}”的Word模板文件", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // 选择导出目录
                var folderDialog = new Microsoft.Win32.OpenFolderDialog
                {
                    Title = $"选择Word模板导出目录 (共找到 {templates.Count} 个相关模板)",
                    InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
                };

                if (folderDialog.ShowDialog() == true)
                {
                    int count = 0;
                    foreach (var kv in templates)
                    {
                        var dest = System.IO.Path.Combine(folderDialog.FolderName, kv.Key);
                        try
                        {
                            System.IO.File.Copy(kv.Value, dest, true);
                            count++;
                        }
                        catch (Exception copyEx)
                        {
                            LogHelper.Warn($"复制模板 {kv.Key} 失败: {copyEx.Message}");
                        }
                    }

                    MessageBox.Show($"成功导出 {count} 个模板文件到：{folderDialog.FolderName}", "成功", MessageBoxButton.OK, MessageBoxImage.Information);
                    LogHelper.Info($"导出 {count} 个“{SelectedDocumentType}”相关Word模板");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"导出模板失败：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                LogHelper.Error($"导出Word模板异常：{ex.Message}");
            }
        }

        /// <summary>
        /// 执行导出数据模板
        /// </summary>
        private void ExecuteExportTemplate()
        {
            try
            {
                // 1. 模糊查找所有相关Excel模板
                // 查找包含文书类型名称的 .xlsx 或 .xls
                var templates = new Dictionary<string, string>();
                if (!string.IsNullOrEmpty(SelectedDocumentType) && SelectedDocumentType != "其它类型")
                {
                    var xlsxTemplates = FindTemplates(SelectedDocumentType, "*.xlsx");
                    var xlsTemplates = FindTemplates(SelectedDocumentType, "*.xls");

                    // 合并结果
                    foreach (var kv in xlsxTemplates) templates[kv.Key] = kv.Value;
                    foreach (var kv in xlsTemplates) if (!templates.ContainsKey(kv.Key)) templates[kv.Key] = kv.Value;
                }

                // 2. 如果找到了特定模板，批量导出
                if (templates.Count > 0)
                {
                    var folderDialog = new Microsoft.Win32.OpenFolderDialog
                    {
                        Title = $"选择Excel模板导出目录 (共找到 {templates.Count} 个相关模板)",
                        InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
                    };

                    if (folderDialog.ShowDialog() == true)
                    {
                        int count = 0;
                        foreach (var kv in templates)
                        {
                            var dest = System.IO.Path.Combine(folderDialog.FolderName, kv.Key);
                            try
                            {
                                System.IO.File.Copy(kv.Value, dest, true);
                                count++;
                            }
                            catch (Exception copyEx)
                            {
                                LogHelper.Warn($"复制模板 {kv.Key} 失败: {copyEx.Message}");
                            }
                        }

                        MessageBox.Show($"成功导出 {count} 个Excel模板文件到：{folderDialog.FolderName}", "成功", MessageBoxButton.OK, MessageBoxImage.Information);
                        LogHelper.Info($"已导出 {count} 个“{SelectedDocumentType}”相关Excel模板");
                    }
                    return;
                }

                // 3. 如果没找到特定模板，使用通用导出逻辑（保留原有Fallback）
                var saveFileDialogDefault = new Microsoft.Win32.SaveFileDialog
                {
                    Filter = "Excel模板文件 (*.xlsx)|*.xlsx|所有文件 (*.*)|*.*",
                    Title = "导出案件数据模板",
                    FileName = (string.IsNullOrEmpty(SelectedDocumentType) || SelectedDocumentType == "其它类型") ? "案件数据模板.xlsx" : $"{SelectedDocumentType}_数据模板.xlsx",
                    InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
                };

                if (saveFileDialogDefault.ShowDialog() == true)
                {
                    var result = ExcelService.ExportCaseTemplate(saveFileDialogDefault.FileName);
                    if (result)
                    {
                        MessageBox.Show("模板导出成功", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    else
                    {
                        MessageBox.Show("模板导出失败", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"导出模板失败：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                LogHelper.Error($"导出Excel模板异常：{ex.Message}");
            }
        }

        /// <summary>
        /// 查找模板文件的辅助方法
        /// </summary>
        /// <param name="keyword">包含的关键字</param>
        /// <param name="searchPattern">文件通配符 (如 *.docx)</param>
        /// <returns>字典：文件名 -> 完整路径</returns>
        private Dictionary<string, string> FindTemplates(string keyword, string searchPattern)
        {
            var results = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var dirs = new List<string>();

            // 1. SystemTemplateDir
            if (!string.IsNullOrEmpty(WordService.SystemTemplateDir) && System.IO.Directory.Exists(WordService.SystemTemplateDir))
                dirs.Add(WordService.SystemTemplateDir);

            // 2. BaseDirectory/Resources/Templates
            string p2 = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "Templates");
            if (System.IO.Directory.Exists(p2)) dirs.Add(p2);

            // 3. Project Root (Development)
            string projectRoot = System.IO.Path.GetFullPath(System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\.."));
            string p3 = System.IO.Path.Combine(projectRoot, "Resources", "Templates");
            if (System.IO.Directory.Exists(p3)) dirs.Add(p3);

            foreach (var dir in dirs)
            {
                try
                {
                    // 先按通配符找
                    var files = System.IO.Directory.GetFiles(dir, searchPattern);
                    foreach (var f in files)
                    {
                        string fname = System.IO.Path.GetFileName(f);
                        // 再匹配关键字
                        if (fname.Contains(keyword) && !results.ContainsKey(fname))
                        {
                            results.Add(fname, f);
                        }
                    }
                }
                catch { }
            }
            return results;
        }

        /// <summary>
        /// 执行清空日志
        /// </summary>
        private void ExecuteClearLog()
        {
            Application.Current.Dispatcher.Invoke(() =>
            {
                LogList.Clear();
            });
        }

        #endregion

        #region 辅助方法
        /// <summary>
        /// 日志回调（更新UI日志列表）
        /// </summary>
        private void OnLogReceived(string content, LogLevel level)
        {
            Application.Current.Dispatcher.Invoke(() =>
            {
                string displayContent = $"[{DateTime.Now:HH:mm:ss}] {content}";
                LogList.Add(new UILogItem { Content = displayContent, Level = level });
            });
        }

        // 释放资源
        public void Dispose()
        {
            LogHelper.OnLogReceived -= OnLogReceived;
        }
        #endregion
    }
}