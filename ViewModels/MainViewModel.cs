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

        #region 日志区绑定属性
        /// <summary>
        /// 日志列表（用于UI展示）
        /// </summary>
        public ObservableCollection<UILogItem> LogList { get; set; } = new ObservableCollection<UILogItem>();
        #endregion

        /// <summary>
        /// 剩余额度显示（格式化）
        /// </summary>
        public string QuotaDisplay => $"剩余额度：{UserInfo.Quota:F2} 元（单价：{UserInfo.UserPrice:F2} 元/条，可处理{(UserInfo.UserPrice > 0 ? (int)(UserInfo.Quota / UserInfo.UserPrice) : 0)}条数据）";


        /// <summary>
        /// 后台消息列表
        /// </summary>
        public ObservableCollection<string> MessageList { get; set; } = new ObservableCollection<string>();
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
        /// 同步额度命令
        /// </summary>
        public ICommand SyncQuotaCommand { get; }

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
            NetworkStatusService.NetworkAvailabilityChanged += OnNetworkChanged;

            // 初始化命令
            LogoutCommand = new RelayCommand(ExecuteLogout);
            RechargeCommand = new RelayCommand(ExecuteRecharge);
            SyncQuotaCommand = new RelayCommand(ExecuteSyncQuota);
            ChangePasswordCommand = new RelayCommand(ExecuteChangePassword);

            // 初始化日志回调（实时更新UI日志）
            LogHelper.OnLogReceived += OnLogReceived;

            // Log welcome message
            LogHelper.Info($"尊敬的{UserInfo.UserName}用户登录成功，欢迎使用法律文书处理工具！");

            // 同步版本信息
            InitVersionInfoAsync();
            // 获取系统公告
            InitSystemMessagesAsync();

            // 首次进入检测网络并同步
            if (NetworkStatusService.IsConnected && UserInfo.Status == UserStatus.OFFLINE)
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
            catch (Exception ex)
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
                catch (Exception ex)
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
                                try { newQuota = d.NewQuota; } catch { } // Optional
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
                        try { waitWindow.Close(); } catch { }
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
            catch (Exception ex)
            {
                LogHelper.Error($"同步额度异常：{ex.Message}");
            }
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

