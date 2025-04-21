using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace MolePCRConvert4WPF.App.Converters
{
    /// <summary>
    /// Converts a boolean value to a Visibility value.
    /// </summary>
    [ValueConversion(typeof(bool), typeof(Visibility))]
    public class BooleanToVisibilityConverter : IValueConverter
    {
        /// <summary>
        /// Gets or sets the Visibility value to return when the input is true.
        /// Defaults to Visible.
        /// </summary>
        public Visibility TrueValue { get; set; } = Visibility.Visible;

        /// <summary>
        /// Gets or sets the Visibility value to return when the input is false.
        /// Defaults to Collapsed.
        /// </summary>
        public Visibility FalseValue { get; set; } = Visibility.Collapsed;

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool boolValue)
            {
                return boolValue ? TrueValue : FalseValue;
            }
            return DependencyProperty.UnsetValue; // Indicate conversion failure
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is Visibility visibilityValue)
            {
                if (visibilityValue == TrueValue)
                    return true;
                if (visibilityValue == FalseValue)
                    return false;
            }
            // If the value is not one of the expected visibilities, 
            // or if TrueValue and FalseValue are the same, it's ambiguous.
            return DependencyProperty.UnsetValue;
        }
    }
} 