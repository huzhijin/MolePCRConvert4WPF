using System;
using System.Globalization;
using System.Windows.Data;

namespace MolePCRConvert4WPF.App.Converters
{
    /// <summary>
    /// 将布尔值转换为字符串的转换器
    /// </summary>
    public class BoolToStringConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool boolValue)
            {
                // 参数格式："TrueString,FalseString,NullString"
                if (parameter is string paramString)
                {
                    string[] parts = paramString.Split(',');
                    if (parts.Length >= 2)
                    {
                        return boolValue ? parts[0] : parts[1];
                    }
                }
                return boolValue.ToString();
            }
            
            // 如果value为null且参数包含第三个部分，则返回该部分
            if (value == null && parameter is string nullParamString)
            {
                string[] parts = nullParamString.Split(',');
                if (parts.Length >= 3)
                {
                    return parts[2];
                }
            }
            
            return string.Empty;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is string stringValue && parameter is string paramString)
            {
                string[] parts = paramString.Split(',');
                if (parts.Length >= 2)
                {
                    if (stringValue == parts[0])
                        return true;
                    if (stringValue == parts[1])
                        return false;
                }
            }
            return false;
        }
    }
} 