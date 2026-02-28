using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using docment_tools_client.Helpers;
using docment_tools_client.Models;
using docment_tools_client.ViewModels;
using docment_tools_client.Views; // 确保引入页面所在命名空间

namespace docment_tools_client.Views
{
    /// <summary>
    /// MainView.xaml 的交互逻辑
    /// </summary>
    public partial class MainView : Window
    {
        public MainView()
        {
            InitializeComponent();

            // 初始化：默认选中第一个菜单（首页）
            if (MenuListBox.Items.Count > 0)
            {
                MenuListBox.SelectedIndex = 0;
            }
        }
       
        /// <summary>
        /// 带用户信息的构造函数
        /// </summary>
        public MainView(UserInfo userInfo) : this()
        {
            DataContext = new MainViewModel(userInfo);
        }

        /// <summary>
        /// 左侧菜单选中变更：跳转对应页面
        /// </summary>
        private void MenuListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // 1. 获取选中的菜单项
            var listBox = sender as ListBox;
            var selectedItem = listBox?.SelectedItem as ListBoxItem;
            if (selectedItem == null || string.IsNullOrEmpty(selectedItem.Tag?.ToString()))
            {
                return;
            }

            // 2. 避免重复加载同一页面（优化体验）
            string targetPageName = selectedItem.Tag.ToString();
            if (MainContentFrame.Content?.GetType().Name == targetPageName)
            {
                return;
            }

            // 3. 反射创建页面实例并导航（传递UserInfo）
            try
            {
                // 拼接完整类型名：命名空间 + 页面类名（必须和实际一致）
                string fullTypeName = $"docment_tools_client.Views.{targetPageName}";
                Type pageType = Type.GetType(fullTypeName);

                if (pageType == null)
                {
                    MessageBox.Show($"未找到页面：{targetPageName}，请检查类名和命名空间！", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // ===== 关键：获取MainViewModel中的UserInfo =====
                var mainViewModel = DataContext as MainViewModel;
                if (mainViewModel == null || mainViewModel.UserInfo == null)
                {
                    MessageBox.Show("当前用户信息为空，无法加载页面！", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                UserInfo userInfo = mainViewModel.UserInfo;

                // 伪造测试用户信息
                //UserInfo userInfo = new UserInfo
                //{
                //    UserId = 1001,
                //    UserName = "测试管理员",
                //    Account = "test_admin",
                //    Token = "TEST_TOKEN_20240520",
                //    Quota = 99999.99m,
                //    UserPrice = 0.001m,
                //    LoginRecordId = 9999,
                //    Status = UserStatus.ONLINE,
                //    EncryptedQuota = "fake_encrypt_quota_test",
                //    EncryptedUserPrice = "fake_encrypt_price_test"
                //};

                // ===== 核心修复：兼容无参/有参构造创建页面 =====
                Page targetPage = null;
                try
                {
                    // 尝试1：无参构造创建（优先）
                    targetPage = (Page)Activator.CreateInstance(pageType);
                }
                catch (MissingMethodException)
                {
                    // 尝试2：调用接收UserInfo参数的构造函数
                    var constructor = pageType.GetConstructor(new[] { typeof(UserInfo) });
                    if (constructor != null)
                    {
                        targetPage = (Page)constructor.Invoke(new object[] { userInfo });
                    }
                    else
                    {
                        // 尝试3：调用其他常见构造（比如ViewModel），若无则抛明确错误
                        throw new InvalidOperationException(
                            $"页面 {targetPageName} 既无无参构造，也无接收 UserInfo 的构造函数！\n" +
                            $"请给页面添加：public {targetPageName}() {{ }}  或  public {targetPageName}(UserInfo userInfo) {{ }}");
                    }
                }

                // ===== 给页面赋值UserInfo（兜底）=====
                var pageUserInfoProp = targetPage.GetType().GetProperty("UserInfo");
                if (pageUserInfoProp != null && pageUserInfoProp.CanWrite)
                {
                    pageUserInfoProp.SetValue(targetPage, userInfo);
                }

                // 导航到目标页面
                MainContentFrame.Navigate(targetPage);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"页面加载失败：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}