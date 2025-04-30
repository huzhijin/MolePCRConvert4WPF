using CommunityToolkit.Mvvm.ComponentModel; // Use CommunityToolkit MVVM features
using CommunityToolkit.Mvvm.Input;
using MolePCRConvert4WPF.Core.Models;
using MolePCRConvert4WPF.Core.Services;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using Microsoft.Extensions.Logging; // Add Logging
using System.Collections.Generic; // For List
using System.ComponentModel; // For INotifyPropertyChanged
using MolePCRConvert4WPF.App.Services; // Assuming IDialogService might be here
using System.Text; // For StringBuilder
using MolePCRConvert4WPF.Core.Services; // Correct namespace for INavigationService

namespace MolePCRConvert4WPF.App.ViewModels
{
    // Represents a unique patient for the right-side list
    public partial class PatientDisplayInfo : ObservableObject
    {
        [ObservableProperty]
        private string? name;
        [ObservableProperty]
        private string? caseNumber;
        [ObservableProperty]
        private string? associatedWells;
    }

    // Make the ViewModel Observable
    public partial class SampleAnalysisViewModel : ObservableObject
    {
        private readonly ILogger<SampleAnalysisViewModel> _logger;
        private readonly IAppStateService _appStateService;
        private readonly INavigationService _navigationService; // Add navigation service
        // private readonly IDialogService _dialogService; // OPTIONAL: Inject a dialog service if you have one

        // Backing field for the wells collection
        [ObservableProperty]
        private ObservableCollection<WellLayout> _wellLayouts = new ObservableCollection<WellLayout>();
        
        // Backing field for the number of columns in the plate for UniformGrid binding
        [ObservableProperty]
        private int _plateColumns = 12; // Default to 12 columns (e.g., 96-well plate)

        // List of unique patients derived from wells
        [ObservableProperty]
        private ObservableCollection<PatientDisplayInfo> _patients = new ObservableCollection<PatientDisplayInfo>();
        
        // Properties for user input
        [ObservableProperty]
        private string? _currentSampleName;
        // [ObservableProperty] // Uncomment if PatientId is added
        // private string? _currentPatientId;
        
        // Properties for the Edit Patient Dialog
        [ObservableProperty]
        private string? _dialogPatientName;
        [ObservableProperty]
        private string? _dialogPatientCaseNumber;
        [ObservableProperty]
        private string? _dialogSelectedWellsDisplay;
        [ObservableProperty]
        private bool _isPatientInfoDialogOpen; // To control dialog visibility
        
        private List<WellLayout> _wellsToEdit = new List<WellLayout>(); // Store wells selected when dialog opens

        // Add ApplyCommand property
        public IRelayCommand ApplyCommand { get; }

        public SampleAnalysisViewModel(ILogger<SampleAnalysisViewModel> logger, 
                                     IAppStateService appStateService, 
                                     INavigationService navigationService /*, IDialogService dialogService */)
        {
            _logger = logger;
            _appStateService = appStateService;
            _navigationService = navigationService; // Assign navigation service
            // _dialogService = dialogService;

            LoadWellData();
            
            // Initialize commands - Explicitly use CommunityToolkit.Mvvm.Input.RelayCommand
            ShowSetPatientInfoDialogCommand = new CommunityToolkit.Mvvm.Input.RelayCommand(ShowSetPatientInfoDialog, CanExecuteOnSelectedWells);
            ConfirmSetPatientInfoCommand = new CommunityToolkit.Mvvm.Input.RelayCommand(ConfirmSetPatientInfo);
            CancelSetPatientInfoCommand = new CommunityToolkit.Mvvm.Input.RelayCommand(CancelSetPatientInfo);
            ClearPatientInfoCommand = new CommunityToolkit.Mvvm.Input.RelayCommand(ClearPatientInfo, CanExecuteOnSelectedWells);
            ClearSelectionCommand = new CommunityToolkit.Mvvm.Input.RelayCommand(ClearSelection);
            SelectAllCommand = new CommunityToolkit.Mvvm.Input.RelayCommand(SelectAll);
            ApplyCommand = new CommunityToolkit.Mvvm.Input.RelayCommand(ExecuteApply, CanExecuteApply); // Initialize ApplyCommand
        }

        private void LoadWellData()
        {
            _logger.LogInformation("Loading well data for Sample Analysis...");
            WellLayouts.Clear(); // Clear previous data
            Patients.Clear();
            var currentPlate = _appStateService.CurrentPlate;

            if (currentPlate?.WellLayouts != null && currentPlate.WellLayouts.Any())
            {
                _logger.LogInformation("Loading plate data from AppStateService. Found {count} well layouts in plate data.",
                    currentPlate.WellLayouts.Count);

                // 保留列数设置
                PlateColumns = currentPlate.Columns;

                int defaultRows = 8; // Assuming standard 8 rows (A-H)
                int defaultColumns = PlateColumns > 0 ? PlateColumns : 12; // Ensure default columns if PlateColumns is 0

                _logger.LogInformation("Creating UI grid with {Rows} rows and {Columns} columns.", defaultRows, defaultColumns);

                // Create the UI grid structure sequentially (A1, A2... H12)
                for (int i = 1; i <= defaultRows; i++) // Rows A-H
                {
                    for (int j = 1; j <= defaultColumns; j++) // Columns 1 to defaultColumns
                    {
                        string rowLetter = ((char)('A' + i - 1)).ToString();
                        string positionString = $"{rowLetter}{j}"; // e.g., "A1"

                        // Try to find the corresponding original well data in AppStateService
                        // Primarily match by Position if available, otherwise Row/Column
                        var existingWell = currentPlate.WellLayouts.FirstOrDefault(w =>
                            (!string.IsNullOrEmpty(w.Position) && w.Position.Equals(positionString, StringComparison.OrdinalIgnoreCase)) ||
                            (string.IsNullOrEmpty(w.Position) && w.Row == rowLetter && w.Column == j) // Fallback if Position is null/empty in source
                        );

                        // Create a NEW WellLayout object for the UI
                        var uiWell = new WellLayout
                        {
                            Row = rowLetter, // Set correct Row for UI position
                            Column = j,      // Set correct Column for UI position
                            // Position should be a calculated property like $"{Row}{Column}" in the WellLayout model
                            IsSelected = false, // Default UI state
                            // Copy data from the original well if found
                            PatientName = existingWell?.PatientName,
                            PatientCaseNumber = existingWell?.PatientCaseNumber,
                            SampleName = existingWell?.SampleName // Copy SampleName too if needed by UI/logic
                        };

                        // Subscribe to PropertyChanged for UI updates
                        uiWell.PropertyChanged += Well_PatientInfoChanged;
                        WellLayouts.Add(uiWell); // Add the NEW UI well to the collection
                    }
                }
                _logger.LogInformation("Created UI well layout with {Count} wells.", WellLayouts.Count);

                // Log count of wells that inherited patient info
                int wellsWithPatient = WellLayouts.Count(w => !string.IsNullOrEmpty(w.PatientName));
                _logger.LogInformation("Found {count} UI wells initialized with patient information from AppStateService.", wellsWithPatient);

                // Update the patient list based on the newly created UI wells
                UpdatePatientsList();
            }
            else
            {
                _logger.LogWarning("No plate data found in AppStateService. Generating default empty 8x{cols} layout.", PlateColumns > 0 ? PlateColumns : 12);
                // Generate a default empty layout if no data exists
                 PlateColumns = PlateColumns > 0 ? PlateColumns : 12; // Ensure PlateColumns has a default
                 int defaultRows = 8;
                 int defaultColumns = PlateColumns;
                 for (int i = 1; i <= defaultRows; i++)
                 {
                     for (int j = 1; j <= defaultColumns; j++)
                     {
                         var emptyWell = new WellLayout
                         {
                             Row = ((char)('A' + i - 1)).ToString(),
                             Column = j,
                             IsSelected = false,
                             PatientName = null,
                             PatientCaseNumber = null,
                             SampleName = null
                         };
                         emptyWell.PropertyChanged += Well_PatientInfoChanged;
                         WellLayouts.Add(emptyWell);
                     }
                 }
                 UpdatePatientsList(); // Update patient list (will be empty)
            }
        }

        // Updates the right-hand side list of unique patients
        private void UpdatePatientsList()
        {
            var patientGroups = WellLayouts
                .Where(w => !string.IsNullOrEmpty(w.PatientName) || !string.IsNullOrEmpty(w.PatientCaseNumber))
                .GroupBy(w => new { Name = w.PatientName ?? string.Empty, Case = w.PatientCaseNumber ?? string.Empty })
                .OrderBy(g => g.Key.Name).ThenBy(g => g.Key.Case)
                .ToList();
                
            Patients.Clear();
            foreach (var group in patientGroups)
            {   
                // Build comma-separated list of wells for this patient
                var wellNames = string.Join(", ", group.Select(w => w.WellName).OrderBy(n => n)); 
                Patients.Add(new PatientDisplayInfo
                {
                    Name = group.Key.Name,
                    CaseNumber = group.Key.Case,
                    AssociatedWells = wellNames
                });
            }
             _logger.LogDebug("Updated Patients list with {Count} unique entries.", Patients.Count);
        }

        // Event handler to update patient list when info changes on a well
        private void Well_PatientInfoChanged(object? sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(WellLayout.PatientName) || e.PropertyName == nameof(WellLayout.PatientCaseNumber))
            {
                // Could be optimized, but regenerating the list is simplest
                Application.Current.Dispatcher.Invoke(UpdatePatientsList); 
            }
            // Also update CanExecute if selection changes
             if (e.PropertyName == nameof(WellLayout.IsSelected))
            {
                 ShowSetPatientInfoDialogCommand.NotifyCanExecuteChanged();
                 ClearPatientInfoCommand.NotifyCanExecuteChanged();
            }
        }

        // Command to clear the selection
        public IRelayCommand ClearSelectionCommand { get; }
        private void ClearSelection()
        {
            _logger.LogInformation("Clearing selection.");
            foreach(var well in WellLayouts)
            {
                 if (well.IsSelected)
                 {
                    well.IsSelected = false;
                 }
            }
        }
        
        // Command to select all wells
        public IRelayCommand SelectAllCommand { get; }
        private void SelectAll()
        {
             _logger.LogInformation("Selecting all wells.");
             foreach(var well in WellLayouts)
             {
                  if (!well.IsSelected)
                  {
                     well.IsSelected = true;
                  }
             }
        }

        // Command to show the Set Patient Info dialog
        public IRelayCommand ShowSetPatientInfoDialogCommand { get; }
        public IRelayCommand ConfirmSetPatientInfoCommand { get; }
        public IRelayCommand CancelSetPatientInfoCommand { get; }
        public IRelayCommand ClearPatientInfoCommand { get; }
        
        private bool CanExecuteOnSelectedWells()
        {
            return WellLayouts.Any(w => w.IsSelected);
        }
        
        private void ShowSetPatientInfoDialog()
        {
            _wellsToEdit = WellLayouts.Where(w => w.IsSelected).ToList();
            if (!_wellsToEdit.Any()) return; // Should be prevented by CanExecute, but double-check
            
            // Pre-fill dialog if all selected wells have the same patient info
            var firstWell = _wellsToEdit.First();
            bool allSame = _wellsToEdit.All(w => w.PatientName == firstWell.PatientName && w.PatientCaseNumber == firstWell.PatientCaseNumber);
            
            DialogPatientName = allSame ? firstWell.PatientName : string.Empty;
            DialogPatientCaseNumber = allSame ? firstWell.PatientCaseNumber : string.Empty;
            
            // Build display string for selected wells
            var sb = new StringBuilder();
            for(int i=0; i < _wellsToEdit.Count; i++)
            {
                 sb.Append(_wellsToEdit[i].WellName);
                 if (i < _wellsToEdit.Count - 1) sb.Append(", ");
                 if (sb.Length > 100) { sb.Append("..."); break; } // Limit display length
            }
            DialogSelectedWellsDisplay = sb.ToString();
            
            IsPatientInfoDialogOpen = true; // Signal the View to show the dialog
             _logger.LogInformation("Showing Set Patient Info dialog for wells: {WellNames}", DialogSelectedWellsDisplay);
            // If using a DialogService: await _dialogService.ShowDialogAsync<PatientInfoDialogViewModel>(this); 
        }
        
        private void ConfirmSetPatientInfo()
        {
             _logger.LogInformation("Confirming Set Patient Info. Name: '{Name}', Case#: '{Case}' for {Count} wells.", 
                                  DialogPatientName, DialogPatientCaseNumber, _wellsToEdit.Count);
            foreach (var well in _wellsToEdit)
            {
                well.PatientName = DialogPatientName;
                well.PatientCaseNumber = DialogPatientCaseNumber;
            }
            // UpdatePatientsList(); // This will be triggered by PropertyChanged handler
            CancelSetPatientInfo(); // Close dialog
        }
        
        private void CancelSetPatientInfo()
        {
             IsPatientInfoDialogOpen = false; // Signal the View to hide the dialog
             _wellsToEdit.Clear();
             _logger.LogDebug("Set Patient Info dialog cancelled/closed.");
        }
        
        private void ClearPatientInfo()
        {
            var wellsToClear = WellLayouts.Where(w => w.IsSelected).ToList();
             _logger.LogInformation("Clearing Patient Info from {Count} selected wells.", wellsToClear.Count);
            foreach (var well in wellsToClear)
            {
                well.PatientName = null;
                well.PatientCaseNumber = null;
            }
            // UpdatePatientsList(); // Triggered by PropertyChanged
        }
        
        // --- Apply Command Logic --- 

        private bool CanExecuteApply()
        {
            // TODO: Add logic to determine if Apply can be executed 
            // (e.g., is there data loaded? Is analysis configured?)
            return true; // Enable by default for now
        }

        private void ExecuteApply()
        {
            _logger.LogInformation("ApplyCommand executed.");

            var targetPlate = _appStateService.CurrentPlate;
            if (targetPlate?.WellLayouts == null || !targetPlate.WellLayouts.Any())
            {
                _logger.LogError("Cannot apply changes: CurrentPlate or its WellLayouts is null or empty in AppStateService.");
                MessageBox.Show("无法应用更改：未找到有效的板数据或孔位布局为空。", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            _logger.LogInformation("Attempting to update patient info in AppStateService...");
            bool changesMade = false;
            int wellsMatchedCount = 0; // Count of UI wells that found at least one match
            int totalUpdatesApplied = 0; // Count of individual original wells updated
            int uiWellsProcessed = 0;
            int uiWellsWithNoMatch = 0;

            // --- ADDED DEBUG LOGGING for Original Wells --- 
            _logger.LogDebug("--- Dumping first 5 original wells from AppStateService BEFORE matching loop: ---");
            int debugCount = 0;
            foreach(var ow_debug in targetPlate.WellLayouts.Take(5))
            {
                debugCount++;
                _logger.LogDebug($"  OriginalWell[{debugCount}] BEFORE Loop: Position='{ow_debug.Position ?? "NULL"}', Row='{ow_debug.Row ?? "NULL"}', Column={ow_debug.Column}");
            }
            _logger.LogDebug("-------------------------------------------------------------------------------");
            // --- END ADDED DEBUG LOGGING ---
            
            // Iterate through UI wells and find matches in the original list
            foreach (var uiWell in this.WellLayouts)
            {
                uiWellsProcessed++;
                string uiPosition = $"{uiWell.Row}{uiWell.Column}"; // Construct position string for the UI well

                // Find ALL original wells that match the current UI well's position
                var matchedOriginalWells = targetPlate.WellLayouts
                                             .Where(originalWell => 
                                                 // Match if Position exists AND matches uiPosition (case-insensitive)
                                                 (!string.IsNullOrEmpty(originalWell.Position) && originalWell.Position.Equals(uiPosition, StringComparison.OrdinalIgnoreCase)) 
                                                 || 
                                                 // OR Match if Row AND Column match uiWell's Row/Column (regardless of originalWell.Position value)
                                                 (originalWell.Row == uiWell.Row && originalWell.Column == uiWell.Column) 
                                             )
                                             .ToList(); // Convert to list to iterate multiple times if needed

                if (matchedOriginalWells.Any())
                {
                    wellsMatchedCount++; // This UI well found at least one match
                    bool updateAppliedForThisUiWell = false;

                    foreach (var originalWell in matchedOriginalWells)
                    {
                        // Update only if different
                        if (originalWell.PatientName != uiWell.PatientName || originalWell.PatientCaseNumber != uiWell.PatientCaseNumber)
                        {
                             _logger.LogDebug("Updating original well at Position '{OrigPos}' (matched UI {UIPos}): Name='{OldName}'->'{NewName}', Case='{OldCase}'->'{NewCase}'",
                                              originalWell.Position ?? $"{originalWell.Row}{originalWell.Column}", // Display original Position if available
                                              uiPosition,
                                              originalWell.PatientName, uiWell.PatientName, 
                                              originalWell.PatientCaseNumber, uiWell.PatientCaseNumber);
                            originalWell.PatientName = uiWell.PatientName;
                            originalWell.PatientCaseNumber = uiWell.PatientCaseNumber;
                            // Update other fields if needed
                            // originalWell.SampleName = uiWell.SampleName; 
                            changesMade = true;
                            updateAppliedForThisUiWell = true;
                            totalUpdatesApplied++;
                        }
                    }
                    // Optional: Log if an update was made for this specific UI well (across potentially multiple original wells)
                    // if (updateAppliedForThisUiWell) _logger.LogTrace("Updates applied originating from UI well {UIPos}", uiPosition);
                }
                else
                {
                    uiWellsWithNoMatch++;
                    // Log only once per run to avoid flooding
                    if (uiWellsWithNoMatch == 1) 
                    { 
                        _logger.LogWarning("Could not find any matching well in AppStateService for UI well {UIPosition}. Subsequent warnings suppressed.", uiPosition);
                        // Log details of the failed UI well for diagnostics
                        _logger.LogDebug("Failed UI Well Details: Row={Row}, Col={Col}", uiWell.Row, uiWell.Column);
                        
                        // --- ADDED DEBUG LOGGING ---
                        _logger.LogDebug("Dumping first 5 original wells from AppStateService for comparison:");
                        int count = 0;
                        foreach(var ow in targetPlate.WellLayouts.Take(5)) // Log first 5 original wells
                        {
                            count++;
                            _logger.LogDebug($"  OriginalWell[{count}]: Position='{ow.Position ?? "NULL"}', Row='{ow.Row ?? "NULL"}', Column={ow.Column}");
                        }
                        // --- END ADDED DEBUG LOGGING ---
                    } 
                }
            }

            _logger.LogInformation("Update process complete. UI Wells Processed: {ProcessedCount}, UI Wells Matched: {MatchedCount}, UI Wells Not Found: {NotFoundCount}, Total Updates Applied to Original Wells: {UpdatesCount}, Changes Made: {ChangesMade}", 
                                 uiWellsProcessed, wellsMatchedCount, uiWellsWithNoMatch, totalUpdatesApplied, changesMade);

            if (uiWellsWithNoMatch > 0)
            {
                 MessageBox.Show($"警告：未能将 {uiWellsWithNoMatch} 个界面孔位与原始数据中的任何孔位进行匹配。这些孔位的患者信息可能未保存。", "匹配警告", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            else if (changesMade)
            {
                _logger.LogInformation("已将用户设置的患者信息更新到 AppStateService.CurrentPlate.WellLayouts 中");
            }
            else
            {
                _logger.LogInformation("No patient information changes detected to apply.");
            }
            
            // Explicitly set back to AppStateService - Keep this for safety.
            _appStateService.CurrentPlate = targetPlate; 
            _logger.LogInformation("Explicitly set the potentially updated CurrentPlate back to AppStateService.");


            // 导航到下一个视图
            _logger.LogInformation("Navigating to PCR Result Analysis View...");
            _navigationService.NavigateTo<PCRResultAnalysisViewModel>();
        }

        // TODO: Implement logic for drag selection and keyboard shortcuts
        // These often involve interaction with the View's code-behind or attached behaviors.
    }
} 