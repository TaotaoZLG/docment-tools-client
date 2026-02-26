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

            // 初始化：默认选中第一个菜单（文书生成），自动加载对应页面
            if (MenuListBox.Items.Count > 0)
            {
                MenuListBox.SelectedIndex = 2;
            }
        }

        /// <summary>
        /// 带用户信息的构造函数
        /// </summary>
        public MainView(UserInfo userInfo) : this()
        {
            DataContext = new ViewModels.MainViewModel(userInfo);
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

                // 创建页面实例
                Page targetPage = (Page)Activator.CreateInstance(pageType);

                // ===== 关键：给页面的ViewModel传递UserInfo =====
                // 方式2（备选）：直接给页面自身赋值UserInfo（若页面有UserInfo属性）
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