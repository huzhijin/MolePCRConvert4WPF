using CommunityToolkit.Mvvm.ComponentModel;
using System;

namespace MolePCRConvert4WPF.Core.Models
{
    /// <summary>
    /// Represents a single row of analysis results for display in the grid.
    /// </summary>
    public partial class AnalysisResultItem : ObservableObject
    {
        // --- Identifying Info ---
        [ObservableProperty]
        private string? _patientName;

        [ObservableProperty]
        private string? _patientCaseNumber; // Store case number if needed for grouping/reporting

        [ObservableProperty]
        private string? _wellPosition; // e.g., "A1"

        [ObservableProperty]
        private string? _channel; // e.g., "FAM", "HEX"

        [ObservableProperty]
        private string? _targetName; // e.g., "GeneX", "InternalControl"

        // --- Raw/Input Data ---
        [ObservableProperty]
        private double? _ctValue;

        [ObservableProperty]
        private string? _ctValueSpecialMark; // e.g., "-", "Undetermined"

        // --- Calculated Results ---
        [ObservableProperty]
        private double? _concentration;

        [ObservableProperty]
        private string? _detectionResult; // e.g., "阳性", "阴性", "未检出", "无效"

        // --- UI Helper Properties ---
        [ObservableProperty]
        private bool _isFirstPatientRow; // Used for visual grouping in the DataGrid

        // Consider adding original WellLayout ID or Plate ID if needed for traceability
        // public Guid WellLayoutId { get; set; }
    }
} 