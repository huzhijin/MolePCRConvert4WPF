using System;
using System.Globalization;
using System.Windows.Data;
using System.Linq;

namespace MolePCRConvert4WPF.App.Converters
{
    /// <summary>
    /// 将空字符串转换为布尔值
    /// </summary>
    public class StringEmptyToBooleanConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is string str)
            {
                return string.IsNullOrWhiteSpace(str);
            }
            return true; // 非字符串视为空
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
} 