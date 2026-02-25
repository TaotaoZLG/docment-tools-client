using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;

namespace docment_tools_client.Views
{
    public partial class FieldSelectionWindow : Window
    {
        public string SelectedHeader { get; private set; } = string.Empty;

        public FieldSelectionWindow(IEnumerable<string> headers)
        {
            InitializeComponent();
            HeadersList.ItemsSource = headers;
        }

        private void HeaderButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Content is string header)
            {
                SelectedHeader = header;
                this.DialogResult = true;
                this.Close();
            }
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }
    }
}