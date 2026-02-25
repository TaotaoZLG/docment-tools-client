using docment_tools_client.Helpers;
using docment_tools_client.ViewModels;
using docment_tools_client.Views;
using System.Configuration;
using System.Data;
using System.Windows;

namespace docment_tools_client
{
    /// <summary>
    /// App.xaml 的交互逻辑
    /// </summary>
    public partial class App : Application
    {
        public App()
        {
            // EPPlus removed, license context not needed
        }

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            try
            {
                // 初始化日志工具
                LogHelper.Info("应用程序启动中...");

                // 检查本地是否有已登录用户（自动登录，可选）
                var userInfo = LocalStorageHelper.GetUserInfo();
                if (userInfo != null)
                {
                    // 解密额度
                    userInfo.Quota = EncryptHelper.DecryptQuota(userInfo.EncryptedQuota);

                    // 打开主页面
                    var mainView = new MainView(userInfo);
                    MainWindow = mainView;
                    mainView.Show();
                    LogHelper.Info("检测到本地已登录用户，直接进入主页面");
                }
                else
                {
                    // 打开登录页面
                    var loginView = new LoginView();
                    loginView.DataContext = new LoginViewModel();
                    MainWindow = loginView;
                    loginView.Show();
                    LogHelper.Info("未检测到本地登录用户，进入登录页面");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"应用启动失败：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                LogHelper.Error($"应用启动异常：{ex.Message}");
                Shutdown();
            }
        }

        protected override void OnExit(ExitEventArgs e)
        {
            base.OnExit(e);
            LogHelper.Info("应用程序已退出");
        }
    }
}
