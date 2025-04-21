using System;
using System.Globalization;
using System.Windows.Data;

namespace MolePCRConvert4WPF.App.Converters
{
    /// <summary>
    /// Converts a boolean value to its inverse.
    /// True becomes False, False becomes True.
    /// </summary>
    [ValueConversion(typeof(bool), typeof(bool))]
    public class InverseBooleanConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool booleanValue)
            {
                return !booleanValue;
            }
            return value; // Return original value if not boolean
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool booleanValue)
            {
                return !booleanValue;
            }
            return value;
        }
    }
} 