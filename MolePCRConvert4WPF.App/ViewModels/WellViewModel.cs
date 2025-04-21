using MolePCRConvert4WPF.Core.Models;
using System.ComponentModel;
using System.Text.RegularExpressions; // For Regex in SampleName parsing

namespace MolePCRConvert4WPF.App.ViewModels
{
    /// <summary>
    /// ViewModel representing a single well in the plate layout for UI binding.
    /// Wraps the WellLayout model.
    /// </summary>
    public class WellViewModel : ViewModelBase
    {
        private readonly WellLayout _wellLayout;
        private bool _isSelected;
        private string? _patientId;
        private string? _patientName; // Added for UI display
        // TODO: Inject or provide access to a patient lookup service

        public WellLayout Model => _wellLayout;

        // --- Properties bound to UI --- Properties bound to UI ---

        public string WellName => _wellLayout.WellName;

        // Underlying model SampleName (WellName_PatientID)
        // Might not be directly displayed/edited in grid, but reflects model state
        public string SampleName => _wellLayout.SampleName ?? string.Empty;

        public string? TargetName => _wellLayout.TargetName;
        public string? Channel => _wellLayout.Channel;
        public double? CtValue => _wellLayout.CtValue;
        public Core.Enums.WellType WellType
        {
            get => _wellLayout.Type;
            set
            {
                if (_wellLayout.Type != value)
                {
                    _wellLayout.Type = value;
                    OnPropertyChanged();
                }
            }
        }

        // UI-specific property for selection
        public bool IsSelected
        {
            get => _isSelected;
            set => SetProperty(ref _isSelected, value);
        }

        // Patient ID (used for applying)
        public string? PatientId
        {
            get => _patientId;
            set
            {
                // Update only if the value actually changes
                if (SetProperty(ref _patientId, value))
                {
                    // Update the underlying model's SampleName according to the rule
                    // Ensure value is not null/empty before appending, handle null case
                    _wellLayout.SampleName = string.IsNullOrWhiteSpace(value)
                                               ? WellName // If ID is cleared, revert SampleName to just WellName
                                               : $"{WellName}_{value}";

                    // Notify that the underlying SampleName has changed
                    OnPropertyChanged(nameof(SampleName));

                    // Trigger Patient Name lookup/update
                    UpdatePatientName(value);
                }
            }
        }

        // Patient Name (for UI display)
        public string? PatientName
        {
            get => _patientName;
            private set => SetProperty(ref _patientName, value); // Private set, updated by lookup
        }

        // --- Constructor --- Constructor ---
        public WellViewModel(WellLayout wellLayout)
        {
            _wellLayout = wellLayout ?? throw new ArgumentNullException(nameof(wellLayout));
            // Initialize PatientId and PatientName based on existing SampleName in the model
            ParseSampleName();
        }

        private void ParseSampleName()
        {
            // Try to extract PatientID from SampleName (e.g., "A1_P001")
            var match = Regex.Match(_wellLayout.SampleName ?? "", @"^.+_([^_]+)$"); // Match text after last underscore
            if (match.Success)
            {
                var parsedId = match.Groups[1].Value;
                _patientId = parsedId; // Set initial _patientId field directly to avoid loop
                UpdatePatientName(parsedId); // Look up name for the parsed ID
            }
            else
            {
                // If no underscore or pattern doesn't match, assume no Patient ID is set
                _patientId = null;
                _patientName = null; // Clear name
            }
            // Notify UI about initial values
            OnPropertyChanged(nameof(PatientId));
            OnPropertyChanged(nameof(PatientName));
        }

        // Placeholder for fetching patient name based on ID
        private async void UpdatePatientName(string? patientId)
        {
            if (string.IsNullOrWhiteSpace(patientId))
            {
                PatientName = null;
                return;
            }

            try
            {
                // --- Placeholder Lookup Logic --- --- Placeholder Lookup Logic ---
                // TODO: Replace with actual call to a patient lookup service
                // Example: _patientName = await _patientLookupService.GetPatientNameByIdAsync(patientId);
                await Task.Delay(10); // Simulate async lookup
                PatientName = $"姓名_{patientId}"; // Dummy name
                // --- End Placeholder --- --- End Placeholder ---
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error looking up patient name for ID {patientId}: {ex.Message}");
                PatientName = "<查询失败>"; // Indicate lookup error
            }
        }
    }
} 