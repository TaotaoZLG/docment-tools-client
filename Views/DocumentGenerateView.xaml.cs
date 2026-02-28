using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using docment_tools_client.Models;
using docment_tools_client.ViewModels;

namespace docment_tools_client.Views
{
    /// <summary>
    /// DocumentGenerateView.xaml 的交互逻辑
    /// </summary>
    public partial class DocumentGenerateView : Page
    {
        public UserInfo UserInfo { get; set; }

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="userInfo">当前登录用户信息（从MainViewModel传递）</param>
        public DocumentGenerateView(UserInfo userInfo)
        {
            InitializeComponent();

            UserInfo = userInfo;

            // 初始化专属VM
            DataContext = new DocumentGenerateViewModel(userInfo);
        }

        // 页面卸载时释放资源
        //private void Page_Unloaded(object sender, RoutedEventArgs e)
        //{
        //    _viewModel.Dispose();
        //}
    }
}
