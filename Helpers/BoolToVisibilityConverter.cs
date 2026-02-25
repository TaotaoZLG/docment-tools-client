using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;
using System.Windows.Markup;

namespace docment_tools_client.Helpers
{
    public class BoolToVisibilityConverter : MarkupExtension, IValueConverter
    {
        public bool IsInverted { get; set; }

        public BoolToVisibilityConverter() { }

        public override object ProvideValue(IServiceProvider serviceProvider)
        {
            return this;
        }

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool boolValue)
            {
                if (IsInverted) boolValue = !boolValue;
                // Check for parameter "Invert" string as well if confused
                if (parameter is string s && s.Equals("Invert", StringComparison.OrdinalIgnoreCase))
                {
                    boolValue = !boolValue;
                }
                
                return boolValue ? Visibility.Visible : Visibility.Collapsed;
            }
            return Visibility.Collapsed;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
