using CommunityToolkit.Mvvm.ComponentModel;
using System.Collections.Generic;
using System.Linq;

namespace MolePCRConvert4WPF.Core.Models
{
    /// <summary>
    /// Maps a PatientInfo object to one or more well positions on the plate.
    /// </summary>
    public partial class PatientWellMapping : ObservableObject
    {
        [ObservableProperty]
        private PatientInfo _patient;

        [ObservableProperty]
        private List<string> _wellPositions; // List of positions like "A1", "H12"

        // Read-only property for display in DataGrid (optional)
        public string WellPositionsDisplay => string.Join(", ", WellPositions ?? Enumerable.Empty<string>());

        public PatientWellMapping(PatientInfo patient, List<string> wellPositions)
        {
            _patient = patient;
            _wellPositions = wellPositions ?? new List<string>();
        }
    }
} 