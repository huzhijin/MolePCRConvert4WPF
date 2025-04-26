using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace MolePCRConvert4WPF.App.Converters
{
    /// <summary>
    /// 当值为0时返回Visible，否则返回Collapsed的转换器
    /// </summary>
    public class ZeroToVisibleConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value == null)
                return Visibility.Visible;
                
            if (value is int intValue)
                return intValue == 0 ? Visibility.Visible : Visibility.Collapsed;
                
            if (value is double doubleValue)
                return doubleValue == 0 ? Visibility.Visible : Visibility.Collapsed;
                
            if (int.TryParse(value.ToString(), out int parsedValue))
                return parsedValue == 0 ? Visibility.Visible : Visibility.Collapsed;
                
            return Visibility.Collapsed;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
} 