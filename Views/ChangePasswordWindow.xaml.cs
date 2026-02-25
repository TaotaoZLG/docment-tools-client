using System;
using System.Windows;
using docment_tools_client.ViewModels;

namespace docment_tools_client.Views
{
    public partial class ChangePasswordWindow : Window
    {
        public ChangePasswordWindow()
        {
            InitializeComponent();
            if (DataContext is ChangePasswordViewModel vm)
            {
                vm.CloseAction = () => this.Close();
            }
        }
    }
}
