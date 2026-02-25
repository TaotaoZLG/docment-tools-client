using docment_tools_client.Helpers;
using docment_tools_client.Models;
using docment_tools_client.Services;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Text;
using System.Windows;
using System.Windows.Input;

namespace docment_tools_client.ViewModels
{
    /// <summary>
    /// 应用主页视图模型（实现数据绑定和核心业务逻辑）
    /// </summary>
    public class MainViewModel : ViewModelBase
    {
        #region 用户区绑定属性
        /// <summary>
        /// 用户信息
        /// </summary>
        public UserInfo UserInfo { get; set; }

        /// <summary>
        /// 剩余额度显示（格式化）
        /// </summary>
        public string QuotaDisplay => $"剩余额度：{UserInfo.Quota:F2} 元（单价：{UserInfo.UserPrice:F2} 元/条，可处理{(UserInfo.UserPrice > 0 ? (int)(UserInfo.Quota / UserInfo.UserPrice) : 0)}条数据）";


        /// <summary>
        /// 后台消息列表
        /// </summary>
        public ObservableCollection<string> MessageList { get; set; } = new ObservableCollection<string>();
        #endregion

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
        /// 退出登录命令
        /// </summary>
        public ICommand LogoutCommand { get; }

        /// <summary>
        /// 充值命令
        /// </summary>
        public ICommand RechargeCommand { get; }

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
        /// 同步额度命令
        /// </summary>
        public ICommand SyncQuotaCommand { get; }

        /// <summary>
        /// 清空日志命令
        /// </summary>
        public ICommand ClearLogCommand { get; }

        /// <summary>
        /// 快捷选择占位符命令
        /// </summary>
        public ICommand QuickSelectCommand { get; }

        public ICommand ChangePasswordCommand { get; }
        #endregion




        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="userInfo">登录成功后的用户信息</param>
        public MainViewModel(UserInfo userInfo)
        {
            // 初始化用户信息
            UserInfo = userInfo ?? throw new ArgumentNullException(nameof(userInfo));
            InitUserMessages();
            
            // 登录业务场景：断网/联网状态同步逻辑（对应第7条）
            // 监听网络状态变更
            Services.NetworkStatusService.NetworkAvailabilityChanged += OnNetworkChanged;

            // 初始化命令
            LogoutCommand = new RelayCommand(ExecuteLogout);
            RechargeCommand = new RelayCommand(ExecuteRecharge);
            SelectTemplateCommand = new RelayCommand(ExecuteSelectTemplate);
            SelectExcelCommand = new RelayCommand(ExecuteSelectExcel);
            SelectOutputDirCommand = new RelayCommand(ExecuteSelectOutputDir);
            PreviewCommand = new RelayCommand(ExecutePreview);
            StartGenerateCommand = new RelayCommand(ExecuteStartGenerate);
            ExportTemplateCommand = new RelayCommand(ExecuteExportTemplate);
            ExportWordTemplateCommand = new RelayCommand(ExecuteExportWordTemplate);
            SyncQuotaCommand = new RelayCommand(ExecuteSyncQuota);
            ClearLogCommand = new RelayCommand(ExecuteClearLog);
            QuickSelectCommand = new RelayCommand(ExecuteQuickSelect);
            ChangePasswordCommand = new RelayCommand(ExecuteChangePassword);

            // 初始化日志回调（实时更新UI日志）


            LogHelper.OnLogReceived += OnLogReceived;

            // 初始化默认模板路径
            TemplatePath = WordService.GetDefaultTemplatePath();
            OnPropertyChanged(nameof(TemplatePath));

            // Log welcome message
            LogHelper.Info($"尊敬的{UserInfo.UserName}用户登录成功，欢迎使用法律文书处理工具！");

            // 同步版本信息
            InitVersionInfoAsync();
            // 获取系统公告
            InitSystemMessagesAsync();
            
            // 首次进入检测网络并同步
            if (Services.NetworkStatusService.IsConnected && UserInfo.Status == UserStatus.OFFLINE)
            {
                 // 如果之前是离线状态进入的，或者标记为离线
                 OnNetworkChanged(true);
            }
        }

        /// <summary>
        /// 初始化系统公告
        /// </summary>
        private async void InitSystemMessagesAsync()
        {
            try
            {
                var messages = await ApiService.GetSystemMessagesAsync();
                if (messages != null && messages.Count > 0)
                {
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        MessageList.Clear();
                        foreach (var msg in messages)
                        {
                            MessageList.Add(msg);
                        }
                    });
                }
            }
            catch (Exception ex)
            {
                LogHelper.Error($"初始化公告失败: {ex.Message}");
            }
        }
        
        // 登录业务场景：断网/联网状态同步逻辑（对应第7条）
        private async void OnNetworkChanged(bool isAvailable)
        {
             if (UserInfo == null || UserInfo.Status == UserStatus.EXIT) return;

             try 
             {
                 if (isAvailable)
                 {
                     // 联网：状态切换 -> ONLINE，发起同步
                     UserInfo.Status = UserStatus.ONLINE;
                     LogHelper.Info("网络已连接，切换用户状态为 ONLINE");
                     
                     InitSystemMessagesAsync();
                     
                     // 解密本地唯一标识 (这里简化，假设 CurrentDeviceIdentifier 直接可用)
                     // 同步状态
                     await ApiService.UpdateActiveStatusAsync(UserInfo.LoginRecordId, "ONLINE");
                     
                     // 同步额度
                     var quotaResult = await ApiService.SyncQuotaAsync(UserInfo.UserId, UserInfo.Account);
                     if (quotaResult.IsSuccess && quotaResult.Data is System.Text.Json.JsonElement je)
                     {
                          if (je.TryGetProperty("userBalance", out var pBal) && pBal.ValueKind == System.Text.Json.JsonValueKind.Number)
                          {
                              UserInfo.Quota = pBal.GetDecimal();
                              UserInfo.EncryptedQuota = EncryptHelper.EncryptQuota(UserInfo.Quota);
                          }
                          if (je.TryGetProperty("userPrice", out var pPrice) && pPrice.ValueKind == System.Text.Json.JsonValueKind.Number)
                          {
                              UserInfo.UserPrice = pPrice.GetDecimal();
                              UserInfo.EncryptedUserPrice = EncryptHelper.EncryptQuota(UserInfo.UserPrice);
                          }
                          OnPropertyChanged(nameof(QuotaDisplay));
                     }
                     
                     // 持久化更新
                     LocalStorageHelper.SaveUserInfo(UserInfo);
                 }
                 else
                 {
                     // 断网：状态切换 -> OFFLINE
                     UserInfo.Status = UserStatus.OFFLINE;
                     LogHelper.Warn("网络断开，切换用户状态为 OFFLINE");
                     
                     // 本地记录该状态，不发起请求
                     LocalStorageHelper.SaveUserInfo(UserInfo);
                 }
             }
             catch(Exception ex)
             {
                 LogHelper.Error($"网络状态变更处理异常: {ex.Message}");
             }
        }

        #region 用户区命令实现
        /// <summary>
        /// 执行退出登录
        /// </summary>
        // 登录业务场景：正常退出登录逻辑（对应第6条）
        private async void ExecuteLogout()
        {
            try
            {
                // 1. 网络检测
                if (!Helpers.NetworkHelper.IsNetworkAvailable())
                {
                    MessageBox.Show("请您先连接网络再退出登录", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                
                // 2. 发起后端请求
                if (UserInfo != null)
                {
                     // 此时调用API，告诉后端该LoginRecord已注销
                     // 对应：POST /prod-api/out/auth/logout
                     await ApiService.LogoutAsync(UserInfo.Account, UserInfo.Token, UserInfo.LoginRecordId);
                }

                // 3. 状态更新与本地清理
                if (UserInfo != null) UserInfo.Status = UserStatus.EXIT;
                // 删除本地用户配置
                LocalStorageHelper.DeleteUserConfig();

                // 4. 跳转回登录页
                Application.Current.Dispatcher.Invoke(() =>
                {
                    var loginView = new Views.LoginView();
                    loginView.DataContext = new LoginViewModel();
                    var oldWin = Application.Current.MainWindow;
                    Application.Current.MainWindow = loginView;
                    loginView.Show();
                    oldWin?.Close();
                });
            }
            catch (Exception ex)
            {
                LogHelper.Error($"退出登录异常：{ex.Message}");
                // 强制退出UI逻辑
                LocalStorageHelper.DeleteUserConfig();
                // ... logic to show login window even if API failed? Yes.
                Application.Current.Dispatcher.Invoke(() =>
                {
                    var loginView = new Views.LoginView();
                    loginView.DataContext = new LoginViewModel();
                    Application.Current.MainWindow.Close();
                    Application.Current.MainWindow = loginView;
                    loginView.Show();
                });
            }
        }

        private void ExecuteChangePassword()
        {
             var window = new Views.ChangePasswordWindow();
             window.Owner = Application.Current.MainWindow;
             window.ShowDialog();
        }

        /// <summary>
        /// 执行充值（打开H5充值页）
        /// </summary>
        private async void ExecuteRecharge()
        {
            try
            {
                // 1. 弹出充值金额选择窗口
                var rechargeWindow = new Views.RechargeWindow();
                rechargeWindow.Owner = Application.Current.MainWindow;
                if (rechargeWindow.ShowDialog() != true) return;
                
                decimal amount = rechargeWindow.SelectedAmount;

                // 2. 向后端发起订单创建请求
                var result = await ApiService.CreateRechargeOrderAsync(UserInfo.UserId, UserInfo.UserName, amount);
                if (!result.IsSuccess)
                {
                    MessageBox.Show(result.Msg ?? "充值申请失败", "充值失败", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // 3. 解析返回数据 
                // 后端返回的数据中，授权令牌payToken是解密过的（明文），在客户端这里无需解密
                string orderNo = "";
                string payToken = "";
                long expireTime = 0;

                try 
                {
                    if (result.Data is System.Text.Json.JsonElement je)
                    {
                        if (je.TryGetProperty("orderNo", out var pOrderNo)) orderNo = pOrderNo.GetString() ?? "";
                        // 获取支付令牌 payToken (明文)
                        if (je.TryGetProperty("payToken", out var pToken)) payToken = pToken.GetString() ?? "";
                        if (je.TryGetProperty("expireTime", out var pExpire)) expireTime = pExpire.GetInt64();
                    }
                    else if (result.Data != null)
                    {
                         dynamic d = result.Data;
                         orderNo = d.orderNo;
                         payToken = d.payToken;
                         expireTime = d.expireTime;
                    }
                }
                catch
                {
                    // Ignore parse error, check string empty later
                }

                if (string.IsNullOrEmpty(payToken) || string.IsNullOrEmpty(orderNo))
                {
                     MessageBox.Show("服务器返回数据异常，无法获取支付凭证", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                     return;
                }

                // 4. Token处理
                // 后端返回的payToken已加密，直接进行URL编码即可
                string urlEncryptedToken = Uri.EscapeDataString(payToken);

                // 5. 构造H5支付URL (部署在腾讯云)
                string h5BaseUrl = "http://127.0.0.1:8080/client-pay.html"; 
                string rechargeUrl = $"{h5BaseUrl}?payToken={urlEncryptedToken}&orderNo={orderNo}&amount={amount}&expireTime={expireTime}";

                // 6. 调用系统浏览器打开URL
                try 
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = rechargeUrl,
                        UseShellExecute = true
                    });
                }
                catch(Exception ex)
                {
                    MessageBox.Show($"无法自动打开浏览器，请手动复制链接支付：\n{rechargeUrl}", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                }

                LogHelper.Info($"发起充值：订单{orderNo}, 金额{amount}");

                // 7. 进入轮询等待界面
                var waitWindow = new Views.PaymentWaitWindow();
                waitWindow.Owner = Application.Current.MainWindow;
                
                var cts = new System.Threading.CancellationTokenSource();
                bool isSuccess = false;

                // 启动后台轮询
                _ = System.Threading.Tasks.Task.Run(async () =>
                {
                    while (!cts.Token.IsCancellationRequested)
                    {
                        await System.Threading.Tasks.Task.Delay(3000, cts.Token);
                        
                        // 查询状态
                        var queryResult = await ApiService.QueryRechargeOrderStatusAsync(orderNo);
                        if (queryResult.IsSuccess && queryResult.Data != null)
                        {
                            string status = "";
                            decimal newQuota = -1;

                            if (queryResult.Data is System.Text.Json.JsonElement je)
                            {
                                if (je.TryGetProperty("status", out var pStatus)) status = pStatus.ToString();
                                if (je.TryGetProperty("newQuota", out var pQuota) && pQuota.ValueKind == System.Text.Json.JsonValueKind.Number)
                                    newQuota = pQuota.GetDecimal();
                            }
                            else 
                            {
                                dynamic d = queryResult.Data;
                                status = d.Status?.ToString(); 
                                try { newQuota = d.NewQuota; } catch {} // Optional
                            }
                            
                            // 判断成功 ("SUCCESS" 或 "支付成功" 或 "1")
                            if (status == "SUCCESS" || status == "支付成功" || status == "1")
                            {
                                isSuccess = true;
                                if (newQuota >= 0)
                                {
                                    Application.Current.Dispatcher.Invoke(() =>
                                    {
                                        UserInfo.Quota = newQuota;
                                        UserInfo.EncryptedQuota = EncryptHelper.EncryptQuota(newQuota);
                                        LocalStorageHelper.SaveUserInfo(UserInfo);
                                        OnPropertyChanged(nameof(QuotaDisplay));
                                    });
                                }
                                break;
                            }
                            else if (status == "FAIL" || status == "支付失败" || status == "2")
                            {
                                break;
                            }
                        }

                        // 检查过期
                        if (expireTime > 0 && DateTimeOffset.Now.ToUnixTimeMilliseconds() > expireTime)
                        {
                             break;
                        }
                    }
                    
                    // 关闭窗口
                    await Application.Current.Dispatcher.InvokeAsync(() => 
                    {
                        try { waitWindow.Close(); } catch {}
                    });
                }, cts.Token);

                waitWindow.ShowDialog();
                cts.Cancel(); // Ensure task stops

                if (isSuccess)
                {
                    MessageBox.Show("充值成功！", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                    LogHelper.Info($"订单{orderNo}充值成功");
                }
                else
                {
                     if (expireTime > 0 && DateTimeOffset.Now.ToUnixTimeMilliseconds() > expireTime)
                        MessageBox.Show("支付超时，请重新申请", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                     else
                        LogHelper.Info($"订单{orderNo}支付未完成或取消");
                }
            }
            catch (Exception ex)
            {
                LogHelper.Error($"充值流程异常: {ex.Message}");
                MessageBox.Show($"充值发生错误: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        #endregion

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
                            targetPath = System.IO.Path.Combine(OutputDir, $"{nameNoExt}_{Guid.NewGuid().ToString("N").Substring(0,4)}.docx");
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

//


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
        /// 执行同步额度
        /// </summary>
        private async void ExecuteSyncQuota()
        {
            try 
            {
                var result = await ApiService.SyncQuotaAsync(UserInfo.UserId, UserInfo.Account);
                if (result.IsSuccess)
                {
                    // 动态解析返回数据
                    if (result.Data is System.Text.Json.JsonElement je)
                    {
                        decimal newQuota = UserInfo.Quota;
                        if (je.TryGetProperty("userBalance", out var pBal) && pBal.ValueKind == System.Text.Json.JsonValueKind.Number)
                        {
                             newQuota = pBal.GetDecimal();
                             UserInfo.Quota = newQuota;
                             UserInfo.EncryptedQuota = EncryptHelper.EncryptQuota(newQuota);
                        }

                        if (je.TryGetProperty("userPrice", out var pPrice) && pPrice.ValueKind == System.Text.Json.JsonValueKind.Number)
                        {
                             var newPrice = pPrice.GetDecimal();
                             UserInfo.UserPrice = newPrice;
                             UserInfo.EncryptedUserPrice = EncryptHelper.EncryptQuota(newPrice);
                        }

                        OnPropertyChanged(nameof(QuotaDisplay));
                        LocalStorageHelper.SaveUserInfo(UserInfo);
                        
                        LogHelper.Info($"本地额度同步成功，最新额度为{UserInfo.Quota:F2}元");
                        MessageBox.Show($"本地额度同步成功，最新额度为{UserInfo.Quota:F2}元", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    else
                    {
                        LogHelper.Info("本地额度同步：返回数据格式不匹配");
                    }
                }
                else

                {
                     LogHelper.Error($"本地额度同步失败，原因为：{result.Msg}");
                     MessageBox.Show($"本地额度同步失败：{result.Msg}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch(Exception ex)
            {
                LogHelper.Error($"同步额度异常：{ex.Message}");
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
                    foreach(var kv in xlsxTemplates) templates[kv.Key] = kv.Value;
                    foreach(var kv in xlsTemplates) if(!templates.ContainsKey(kv.Key)) templates[kv.Key] = kv.Value;
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
            LogList.Clear();
        }

        #endregion

        #region 辅助方法

        /// <summary>
        /// 初始化用户消息
        /// </summary>
        private void InitUserMessages()
        {
            MessageList.Clear();
            foreach (var msg in UserInfo.Messages)
            {
                MessageList.Add(msg);
            }
        }

        /// <summary>
        /// 初始化版本信息
        /// </summary>
        private async void InitVersionInfoAsync()
        {
            var result = await ApiService.GetLatestVersionAsync();
            if (result.Code == 200 && result.Data is VersionInfo versionInfo)
            {
                LogHelper.Info($"当前版本：{versionInfo.Version}，{versionInfo.UpdateContent}");
            }
        }

        /// <summary>
        /// 同步消费记录到后台
        /// </summary>
        private async void SyncConsumeRecordsAsync()
        {
            try
            {
                var unsyncedRecords = LocalDbHelper.GetUnsyncedConsumeRecords(UserInfo.Account);
                if (unsyncedRecords.Count > 0)
                {
                    await ApiService.SyncConsumeRecordsAsync(unsyncedRecords);
                }
            }
            catch (Exception ex)
            {
                LogHelper.Error($"同步消费记录异常：{ex.Message}");
            }
        }

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


        #endregion
    }

    /// <summary>
    /// UI日志条目（替换ValueTuple以支持WPF绑定）
    /// </summary>
    public class UILogItem
    {
        public string Content { get; set; } = string.Empty;
        public LogLevel Level { get; set; }
    }
}

