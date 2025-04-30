using MolePCRConvert4WPF.App.Extensions;
// using MolePCRConvert4WPF.Core.Models; // Removed or commented out if Core.PCRResultItem is not needed here
using MolePCRConvert4WPF.App.ViewModels; // Add using for ViewModels namespace
using System;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Media;

namespace MolePCRConvert4WPF.App.Views
{
    /// <summary>
    /// 将PCRResultItem转换为前景色的转换器
    /// </summary>
    public class ResultForegroundConverter : IValueConverter
    {
        /// <summary>
        /// 将PCRResultItem转换为前景色
        /// </summary>
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            // Explicitly use the ViewModel's PCRResultItem
            if (value is ViewModels.PCRResultItem item)
            {
                // Assuming PCRResultItemExtensions is updated or also uses ViewModels.PCRResultItem
                return PCRResultItemExtensions.GetResultForeground(item);
            }
            
            return Brushes.Black;
        }

        /// <summary>
        /// 将前景色转换为PCRResultItem（不实现）
        /// </summary>
        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
} 