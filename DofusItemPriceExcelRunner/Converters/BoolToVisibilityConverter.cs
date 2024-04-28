using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace DofusItemPriceExcelRunner.Converters
{
    internal class BoolToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            bool useHidden = false;
            if(bool.TryParse(parameter?.ToString(), out bool paramBool))
            {
                useHidden = paramBool;
            }
            Visibility valIfNotVisible = useHidden ? Visibility.Hidden : Visibility.Collapsed;
            return value is bool valBool && valBool
                ? Visibility.Visible
                : valIfNotVisible;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return 0;
        }
    }
}
