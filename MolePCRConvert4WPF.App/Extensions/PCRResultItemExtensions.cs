using MolePCRConvert4WPF.Core.Models;
using System.Windows.Media;

namespace MolePCRConvert4WPF.App.Extensions
{
    /// <summary>
    /// PCRResultItem的UI扩展
    /// </summary>
    public static class PCRResultItemExtensions
    {
        /// <summary>
        /// 获取结果文本颜色
        /// </summary>
        /// <param name="item">PCR结果项</param>
        /// <returns>文本颜色</returns>
        public static Brush GetResultForeground(this PCRResultItem item)
        {
            return item.IsPositive ? Brushes.Red : Brushes.Black;
        }
    }
} 