using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows.Data;

namespace MolePCRConvert4WPF.App.Converters
{
    /// <summary>
    /// Converts an integer count into a sequence of numbers starting from a specified value.
    /// Used for generating row/column headers dynamically.
    /// Example: Count=12, Start=1 -> { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12 }
    /// </summary>
    public class SequenceConverter : IValueConverter
    {
        /// <summary>
        /// The starting number of the sequence (default is 1).
        /// </summary>
        public int Start { get; set; } = 1;

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is int count && count > 0)
            {
                // Generate a sequence from Start to Start + count - 1
                return Enumerable.Range(Start, count);
            }

            // Return an empty sequence if the input is invalid
            return Enumerable.Empty<int>();
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            // Conversion back is not supported
            throw new NotSupportedException();
        }
    }
} 