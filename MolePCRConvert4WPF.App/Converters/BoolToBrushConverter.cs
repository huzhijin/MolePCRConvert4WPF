using System;
using System.Globalization;
using System.Windows; // Required for DependencyObject and DependencyProperty
using System.Windows.Data;
using System.Windows.Media;

namespace MolePCRConvert4WPF.App.Converters {
    /// <summary>
    /// 布尔值到画刷的转换器
    /// </summary>
    public class BoolToBrushConverter : DependencyObject, IValueConverter { // Inherit from DependencyObject

        // DependencyProperty for TrueBrush
        public static readonly DependencyProperty TrueBrushProperty =
            DependencyProperty.Register("TrueBrush", typeof(Brush), typeof(BoolToBrushConverter), new PropertyMetadata(null));

        /// <summary>
        /// 为 True 时使用的画刷
        /// </summary>
        public Brush? TrueBrush {
            get { return (Brush?)GetValue(TrueBrushProperty); }
            set { SetValue(TrueBrushProperty, value); }
        }

        // DependencyProperty for FalseBrush
        public static readonly DependencyProperty FalseBrushProperty =
            DependencyProperty.Register("FalseBrush", typeof(Brush), typeof(BoolToBrushConverter), new PropertyMetadata(null));

        /// <summary>
        /// 为 False 时使用的画刷
        /// </summary>
        public Brush? FalseBrush {
            get { return (Brush?)GetValue(FalseBrushProperty); }
            set { SetValue(FalseBrushProperty, value); }
        }

        /// <summary>
        /// 转换方法
        /// </summary>
        /// <param name="value">布尔值</param>
        /// <param name="targetType">目标类型</param>
        /// <param name="parameter">参数</param>
        /// <param name="culture">文化信息</param>
        /// <returns>对应的画刷</returns>
        public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture) {
            if (value is bool isTrue) {
                return isTrue ? TrueBrush : FalseBrush;
            }
            // Return FalseBrush if value is not bool or if it's null
            return FalseBrush;
        }

        /// <summary>
        /// 反向转换方法 (通常不需要为这种转换器实现)
        /// </summary>
        public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture) {
             // Check if the incoming brush matches TrueBrush or FalseBrush
             if (value is Brush brush)
             {
                 if (object.Equals(brush, TrueBrush)) return true;
                 if (object.Equals(brush, FalseBrush)) return false;
             }
             // Return DependencyProperty.UnsetValue if the brush cannot be converted back
             return DependencyProperty.UnsetValue;
            // throw new NotImplementedException(); // Original implementation
        }
    }
} 