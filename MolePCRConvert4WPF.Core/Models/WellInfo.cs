using CommunityToolkit.Mvvm.ComponentModel;

namespace MolePCRConvert4WPF.Core.Models
{
    /// <summary>
    /// Represents a single well on the plate for sample naming purposes.
    /// Inherits from ObservableObject for potential future binding needs within the cell itself.
    /// </summary>
    public partial class WellInfo : ObservableObject
    {
        [ObservableProperty]
        private string? _row; // e.g., "A", "B", ...

        [ObservableProperty]
        private int _column; // 1-based column index (1, 2, 3, etc.)

        public string Position => $"{Row}{Column}";

        [ObservableProperty]
        private string? _sampleName;

        // Add other properties if needed from the original Plate's WellLayout, like WellType
        // Example:
        // [ObservableProperty]
        // private WellType _wellType = WellType.Unknown; 
        // Needs WellType enum defined

        public WellInfo(string row, int column, string? sampleName = null)
        {
            _row = row;
            _column = column;
            _sampleName = sampleName ?? string.Empty;
        }
    }
} 