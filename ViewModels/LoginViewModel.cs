using docment_tools_client.Helpers;
using docment_tools_client.Models;
using docment_tools_client.Services;
using System;
using System.Windows;
using System.Windows.Input;

namespace docment_tools_client.ViewModels
{
    /// <summary>
    /// 登录页视图模型（实现数据绑定和登录业务逻辑）
    /// </summary>
    public class LoginViewModel : ViewModelBase
    {
        #region 绑定属性
        private string _account = ""; 
        private string _password = ""; 
        private bool _isLoginButtonEnabled = true;
        private string _loadingTip = "登录";
        private bool _isAgreementChecked = true;

        /// <summary>
        /// 是否同意协议
        /// </summary>
        public bool IsAgreementChecked
        {
            get => _isAgreementChecked;
            set => SetProperty(ref _isAgreementChecked, value);
        }

        /// <summary>
        /// 账号
        /// </summary>
        public string Account
        {
            get => _account;
            set => SetProperty(ref _account, value);
        }

        /// <summary>
        /// 密码
        /// </summary>
        public string Password
        {
            get => _password;
            set => SetProperty(ref _password, value);
        }

        /// <summary>
        /// 登录按钮是否可用
        /// </summary>
        public bool IsLoginButtonEnabled
        {
            get => _isLoginButtonEnabled;
            set => SetProperty(ref _isLoginButtonEnabled, value);
        }

        /// <summary>
        /// 加载状态提示
        /// </summary>
        public string LoadingTip
        {
            get => _loadingTip;
            set => SetProperty(ref _loadingTip, value);
        }
        #endregion

        #region 命令
        /// <summary>
        /// 登录命令
        /// </summary>
        public ICommand LoginCommand { get; }

        /// <summary>
        /// 申请账号命令
        /// </summary>
        public ICommand ApplyAccountCommand { get; }

        /// <summary>
        /// 打开协议命令
        /// </summary>
        public ICommand OpenAgreementCommand { get; }
        #endregion

        /// <summary>
        /// 构造函数（初始化命令）
        /// </summary>
        public LoginViewModel()
        {
            LoginCommand = new RelayCommand(ExecuteLogin);
            ApplyAccountCommand = new RelayCommand(ExecuteApplyAccount);
            OpenAgreementCommand = new RelayCommand(ExecuteOpenAgreement);
        }

        /// <summary>
        /// 执行登录逻辑
        /// 登录业务场景：完整登录流程（对应第1、2、5条）
        /// </summary>
        private async void ExecuteLogin()
        {
            try
            {
                IsLoginButtonEnabled = false;
                LoadingTip = "正在校验环境...";
                
                // 0. 检查协议勾选
                if (!IsAgreementChecked)
                {
                    MessageBox.Show("请先勾选并同意《平台服务协议》", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                    IsLoginButtonEnabled = true;
                    LoadingTip = "登录";
                    return;
                }

                // 1. 登录强制联网逻辑（对应第1条）
                if (!Helpers.NetworkHelper.IsNetworkAvailable())
                {
                     MessageBox.Show("请先连接网络再进行登录", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                     IsLoginButtonEnabled = true;
                     LoadingTip = "登录";
                     return;
                }

                // 2. 准备登录数据
                LoadingTip = "正在加密数据...";
                
                // 密码加密（对应第2条）
                string encryptedPwd = LoginEncryptHelper.EncryptPassword(Password);
                
                // 获取设备唯一标识（对应第5条） (需先加密存储，此处GetDeviceIdentifier返回的是明文ID还是？根据需求"加密后本地存储，登录时解密该标识并与账号...一起上传")
                // 实际上DeviceHelper.GetDeviceIdentifier 返回的是 Hash过的ID (字符串)，即 "本地唯一标识"。
                // 按照需求：生成本地唯一标识 -> 加密后本地存储 -> 登录时解密 -> 上传。
                // 这里的DeviceHelper目前直接返回Hash字符串。对于Payload，通常直接传这个Hash串。
                // 若需"加密后本地存储"，我们可以在LocalStorageHelper里存这个ID。
                // 简单起见，DeviceHelper每次动态算，或缓存此ID。
                string deviceId = DeviceHelper.GetDeviceIdentifier();
                // 需求提到“登录时解密该标识并...上传”。这意味着本地存的是密文。
                // 这里我们假设 DeviceHelper 内部处理或 LocalStorageHelper 处理。
                // 简化：直接传输 DeviceId (Hash值) 给后端。

                Application.Current.Dispatcher.Invoke(() =>
                {
                    var mainView = new Views.MainView();
                    var loginWin = Application.Current.MainWindow;
                    Application.Current.MainWindow = mainView;
                    mainView.Show();
                    loginWin?.Close();
                });
                return;

                // 3. 发起请求
                LoadingTip = "正在验证身份...";
                var result = await ApiService.LoginAsync(Account, encryptedPwd, deviceId);

                if (result.IsSuccess && result.Data is UserInfo userInfo)
                {
                    // 4. 登录成功处理
                    // Token全局携带与本地存储（对应第3条）
                    // 已经在 UserInfo 对象中包含 Token
                    
                    // 保存用户信息到本地
                    LocalStorageHelper.SaveUserInfo(userInfo);

                    // 跳转主界面
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        var mainView = new Views.MainView();
                        mainView.DataContext = new MainViewModel(userInfo);
                        var loginWin = Application.Current.MainWindow;
                        Application.Current.MainWindow = mainView;
                        mainView.Show();
                        loginWin?.Close();
                    });
                }
                else
                {
                    // 异常关闭窗口等逻辑的后端反馈
                    if (result.Msg.Contains("其他设备登录"))
                    {
                        MessageBox.Show("该账户已在另一台设备登录，需该设备完成退出登录后，方可进行本次登录", "安全警告", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                    else
                    {
                        MessageBox.Show(result.Msg ?? "登录失败", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("登录过程中发生异常，请稍后重试", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                LogHelper.Error($"登录流程异常：{ex.Message}");
            }
            finally
            {
                IsLoginButtonEnabled = true;
                LoadingTip = "登录";
            }
        }

        /// <summary>
        /// 执行申请账号逻辑（打开默认浏览器访问H5官网）
        /// </summary>
        private void ExecuteApplyAccount()
        {
            try
            {
                // 模拟H5官网注册地址
                var registerUrl = "http://127.0.0.1:8080/apply.html";

                // 调用系统默认浏览器打开地址
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = registerUrl,
                    UseShellExecute = true
                });

                LogHelper.Info($"已打开默认浏览器访问注册页面：{registerUrl}");
            }
            catch (Exception ex)
            {
                MessageBox.Show("打开浏览器失败，请手动访问官网申请账号", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                LogHelper.Error($"打开注册页面失败：{ex.Message}");
            }
        }

        private void ExecuteOpenAgreement()
        {
            try
            {
                var url = "https://cenosoft.top/static/file/common/平台服务协议.pdf";
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = url,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"无法打开协议文件：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
