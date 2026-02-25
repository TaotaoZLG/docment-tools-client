using docment_tools_client.Models;
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

            // 初始化默认选中第一个菜单，显示文书生成页
            //MenuListBox.SelectedIndex = 0;
        }

        /// <summary>
        /// 带用户信息的构造函数
        /// </summary>
        public MainView(UserInfo userInfo) : this()
        {
            DataContext = new ViewModels.MainViewModel(userInfo);
        }

        /// <summary>
        /// 菜单选择变更：切换右侧页面
        /// </summary>
        private void MenuListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (MenuListBox.SelectedItem is not ListBoxItem selectedItem) return;

            var pageTag = selectedItem.Tag.ToString();
            var pageUri = new Uri($"{pageTag}.xaml", UriKind.Relative);

            // 避免重复导航同一页面
            if (MainContentFrame.Source != pageUri)
            {
                MainContentFrame.Source = pageUri;
            }
        }
    }
}
