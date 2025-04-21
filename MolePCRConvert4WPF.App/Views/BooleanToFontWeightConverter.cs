using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace MolePCRConvert4WPF.App.Views
{
    /// <summary>
    /// 布尔值转字体粗细转换器
    /// </summary>
    public class BooleanToFontWeightConverter : IValueConverter
    {
        /// <summary>
        /// 将布尔值转换为字体粗细
        /// </summary>
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool boolValue && boolValue)
            {
                return FontWeights.Bold;
            }
            
            return FontWeights.Normal;
        }

        /// <summary>
        /// 将字体粗细转换为布尔值（不实现）
        /// </summary>
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
} 