using System.Windows;

namespace docment_tools_client.Views
{
    public partial class PaymentWaitWindow : Window
    {
        public PaymentWaitWindow()
        {
            InitializeComponent();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}