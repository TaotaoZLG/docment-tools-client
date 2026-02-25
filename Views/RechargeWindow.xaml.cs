using System.Windows;
using System.Text.RegularExpressions;
using System.Windows.Input;

namespace docment_tools_client.Views
{
    public partial class RechargeWindow : Window
    {
        public decimal SelectedAmount { get; private set; } = 0;

        public RechargeWindow()
        {
            InitializeComponent();
        }

        private void Amount_Checked(object sender, RoutedEventArgs e)
        {
            if (sender is System.Windows.Controls.RadioButton rb && rb.Tag != null)
            {
                if (int.TryParse(rb.Tag.ToString(), out int val))
                {
                    // 添加空值检查，防止在 InitializeComponent 期间崩溃
                    if (CustomAmountPanel != null)
                    {
                        CustomAmountPanel.Visibility = (val == 0) ? Visibility.Visible : Visibility.Collapsed;
                    }

                    // 即使控件未加载完全，也可以更新数据属性
                    SelectedAmount = (val == 0) ? 0 : val;
                }
            }
        }

        private void ConfirmButton_Click(object sender, RoutedEventArgs e)
        {
            // If custom amount is visible, parse it
            if (CustomAmountPanel.Visibility == Visibility.Visible)
            {
                if (decimal.TryParse(CustomAmountBox.Text, out decimal customVal) && customVal > 0)
                {
                    SelectedAmount = customVal;
                }
                else
                {
                     MessageBox.Show("请输入有效的金额", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                     return;
                }
            }

            if (SelectedAmount <= 0)
            {
                MessageBox.Show("请选择或输入有效的充值金额", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            this.DialogResult = true;
            this.Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }

        private void NumberValidationTextBox(object sender, TextCompositionEventArgs e)
        {
            Regex regex = new Regex("[^0-9.]+");
            e.Handled = regex.IsMatch(e.Text);
        }
    }
}