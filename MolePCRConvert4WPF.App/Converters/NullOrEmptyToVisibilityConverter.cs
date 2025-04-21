using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace MolePCRConvert4WPF.App.Converters // Adjust namespace if needed
{
    /// <summary>
    /// Converts a null or empty string to Visibility.Collapsed, otherwise Visibility.Visible.
    /// </summary>
    public class NullOrEmptyToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            return string.IsNullOrEmpty(value as string) ? Visibility.Collapsed : Visibility.Visible;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            // ConvertBack is not typically needed for this scenario
            throw new NotImplementedException();
        }
    }
} 