using System;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Markup;

namespace docment_tools_client.Helpers
{
    public class InverseBoolConverter : MarkupExtension, IValueConverter
    {
        public InverseBoolConverter() { }

        public override object ProvideValue(IServiceProvider serviceProvider)
        {
            return this;
        }

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool boolValue)
            {
                return !boolValue;
            }
            return false;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
             if (value is bool boolValue)
            {
                return !boolValue;
            }
            return false;
        }
    }
}
