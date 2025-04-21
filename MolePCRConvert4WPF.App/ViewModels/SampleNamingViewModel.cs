using System;
using System.Collections.Generic; // Added for List
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows.Input; // For ICommand
using MolePCRConvert4WPF.Core.Models;
using Microsoft.Extensions.DependencyInjection; // For IServiceProvider
using System.Windows; // For Application.Current, MessageBox
using Microsoft.Extensions.Logging; // Added for ILogger
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using System.Threading.Tasks;
using MolePCRConvert4WPF.Core.Services; // Added for IDialogService
using MolePCRConvert4WPF.App.ViewModels.Dialogs; // Added for PatientInfoDialogViewModel
using MolePCRConvert4WPF.App.ViewModels; // For PCRResultAnalysisViewModel

namespace MolePCRConvert4WPF.App.ViewModels
{
    public partial class SampleNamingViewModel : ObservableObject
    {
        private readonly ILogger<SampleNamingViewModel> _logger;
        private readonly IDialogService _dialogService; // Injected Dialog Service
        private readonly IAppStateService _appStateService; // Service to access shared Plate data
        private readonly INavigationService _navigationService; // Service for navigation

        [ObservableProperty]
        private ObservableCollection<WellInfo> _wells;

        [ObservableProperty]
        private ObservableCollection<PatientWellMapping> _patientMappings;

        // Property to hold the currently selected wells from the ListView/ItemsControl
        // This needs careful handling in the View's code-behind or via attached behaviors 
        // as direct binding to SelectedItems is tricky.
        // For now, commands will likely operate on a parameter passed from the View.
        // private ObservableCollection<WellInfo> _selectedWells;

        public SampleNamingViewModel(
            ILogger<SampleNamingViewModel> logger, 
            IDialogService dialogService, 
            IAppStateService appStateService,
            INavigationService navigationService)
        {
            _logger = logger;
            _dialogService = dialogService; // Store injected service
            _appStateService = appStateService; // Store injected service
            _navigationService = navigationService; // Store injected service
            _wells = new ObservableCollection<WellInfo>();
            _patientMappings = new ObservableCollection<PatientWellMapping>();

            // Initialize or load data
            LoadDataFromAppState(); // Load existing well/patient info if available
            if (!Wells.Any())
            {
                InitializePlateLayout(); // Initialize if app state had no plate
            }

            // Initialize commands
            SetPatientInfoCommand = new AsyncRelayCommand<IList<object>>(SetPatientInfoAsync, CanSetPatientInfo); // Expects selected WellInfo objects
            RemovePatientMappingCommand = new CommunityToolkit.Mvvm.Input.RelayCommand<PatientWellMapping>(RemovePatientMapping, CanRemovePatientMapping);
            ClearAllPatientsCommand = new CommunityToolkit.Mvvm.Input.RelayCommand(ClearAllPatients, CanClearAllPatients);
            ApplyCommand = new CommunityToolkit.Mvvm.Input.RelayCommand(ApplySampleNames, CanApplySampleNames); // Add CanExecute if needed
        }

        // Load data when the view/VM is activated
        private void LoadDataFromAppState()
        {
            var currentPlate = _appStateService.CurrentPlate;
            if (currentPlate?.WellLayouts != null)
            {
                 _logger.LogInformation("Loading plate data from AppState. Plate ID: {PlateId}", currentPlate.Id);
                 Wells.Clear();
                 PatientMappings.Clear();
                 var tempMappings = new Dictionary<string, PatientWellMapping>(); // Temp store to group wells by patient

                 foreach(var plateWell in currentPlate.WellLayouts)
                 {
                     var wellInfo = new WellInfo(plateWell.Row, plateWell.Column, plateWell.SampleName);
                     Wells.Add(wellInfo);
                     
                     // --- Simplified Reconstruction Logic --- 
                     // Reconstruct PatientMappings based only on non-empty SampleName
                     if (!string.IsNullOrEmpty(plateWell.SampleName))
                     {
                         // Use SampleName as the key for grouping for simplicity. 
                         // A more robust key might involve MRN if available via Sample object later.
                         string patientKey = plateWell.SampleName; 
                         if (!tempMappings.TryGetValue(patientKey, out var mapping))
                         {
                             // Create a placeholder PatientInfo using the SampleName.
                             // You might need to fetch/create more complete PatientInfo later.
                             var patientPlaceholder = new PatientInfo(plateWell.SampleName, null); // MRN is unknown here
                             mapping = new PatientWellMapping(patientPlaceholder, new List<string>());
                             tempMappings.Add(patientKey, mapping);
                         }
                         mapping.WellPositions.Add(wellInfo.Position);
                     }
                     // --- End Simplified Reconstruction --- 
                 }
                 // Add reconstructed mappings to the observable collection
                 foreach(var mapping in tempMappings.Values) PatientMappings.Add(mapping);
                 _logger.LogInformation("Loaded {WellCount} wells and reconstructed {MappingCount} patient mappings.", Wells.Count, PatientMappings.Count);
            }
            else
            {
                 _logger.LogWarning("AppStateService.CurrentPlate is null or has no WellLayouts.");
            }
        }

        private void InitializePlateLayout(int rows = 8, int columns = 12)
        {
            _logger.LogInformation("Initializing plate layout ({Rows}x{Columns}).", rows, columns);
            Wells.Clear();
            PatientMappings.Clear(); // Also clear mappings when re-initializing layout

            char[] rowLabels = { 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H' }; // Assuming max 8 rows
            if (rows > rowLabels.Length)
            {
                 _logger.LogWarning("Requested rows ({Rows}) exceed available labels ({LabelCount}). Clamping to {LabelCount}.", rows, rowLabels.Length, rowLabels.Length);
                 rows = rowLabels.Length;
            }
            
            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < columns; c++) 
                {
                    Wells.Add(new WellInfo(rowLabels[r].ToString(), c + 1));
                }
            }
             _logger.LogInformation("Plate layout initialized with {WellCount} wells.", Wells.Count);
        }

        // --- Commands --- 

        public IAsyncRelayCommand<IList<object>> SetPatientInfoCommand { get; }
        public ICommand RemovePatientMappingCommand { get; }
        public ICommand ClearAllPatientsCommand { get; }
        public ICommand ApplyCommand { get; }

        private bool CanSetPatientInfo(IList<object>? selectedItems)
        {
            // Enable if one or more wells are selected
            return selectedItems != null && selectedItems.Count > 0;
        }

        private async Task SetPatientInfoAsync(IList<object>? selectedItems)
        {
            if (!CanSetPatientInfo(selectedItems)) return;

            var selectedWells = selectedItems!.Cast<WellInfo>().ToList();
            var selectedPositions = selectedWells.Select(w => w.Position).ToList();
            _logger.LogInformation("SetPatientInfoCommand executed for wells: {WellPositions}", string.Join(", ", selectedPositions));

            // --- Prepare Dialog --- 
            // Check if editing existing patient info (based on first selected well)
            PatientInfo? existingPatient = null;
            string? initialName = null;
            string? initialMrn = null;
            var firstWellMapping = PatientMappings.FirstOrDefault(m => m.WellPositions.Contains(selectedWells.First().Position));
            if (firstWellMapping != null)
            {
                existingPatient = firstWellMapping.Patient;
                initialName = existingPatient.Name;
                initialMrn = existingPatient.MedicalRecordNumber;
                _logger.LogDebug("Pre-filling dialog for existing patient: Name={Name}, MRN={Mrn}", initialName, initialMrn);
            }

            var dialogViewModel = new PatientInfoDialogViewModel(initialName, initialMrn);

            // --- Show Dialog --- 
            _logger.LogDebug("Showing Patient Info Dialog...");
            var result = await _dialogService.ShowDialogAsync(dialogViewModel);
            _logger.LogDebug("Dialog closed. Result type: {ResultType}", result?.GetType().Name ?? "null");


            // --- Process Dialog Result --- 
            if (result is PatientInfoDialogViewModel returnedVm) // Check if OK was pressed (dialog returns the ViewModel)
            { 
                string? name = returnedVm.PatientName?.Trim();
                string? mrn = returnedVm.MedicalRecordNumber?.Trim();

                if (string.IsNullOrWhiteSpace(name) && string.IsNullOrWhiteSpace(mrn))
                {
                    _logger.LogWarning("Dialog returned OK but both Name and MRN are empty. Treating as cancel.");
                    return; // Or treat as clearing the patient info?
                }

                // Use the potentially updated existing patient or create a new one
                PatientInfo patientToApply = existingPatient ?? new PatientInfo();
                patientToApply.Name = name; 
                patientToApply.MedicalRecordNumber = mrn;
                
                _logger.LogInformation("Dialog confirmed. Applying Patient Info: Name={PatientName}, MRN={PatientMRN}", patientToApply.Name, patientToApply.MedicalRecordNumber);

                // --- Apply Logic (same as before, but using patientToApply) --- 

                // Remove existing mappings for the selected wells first
                var existingMappings = PatientMappings.Where(m => m.WellPositions.Any(wp => selectedPositions.Contains(wp))).ToList();
                foreach (var existing in existingMappings)
                {
                    var remainingPositions = existing.WellPositions.Except(selectedPositions).ToList();
                    if (remainingPositions.Count > 0)
                    {
                        existing.WellPositions = remainingPositions; 
                         _logger.LogDebug("Updated existing mapping for patient {PatientName}. Remaining wells: {Wells}", existing.Patient.Name, string.Join(",", remainingPositions));
                    }
                    else
                    {
                        PatientMappings.Remove(existing); 
                         _logger.LogDebug("Removed empty existing mapping for patient {PatientName}", existing.Patient.Name);
                    }
                }
                
                // Check if this patient info already exists in other mappings (after potential updates)
                var mappingForThisPatient = PatientMappings.FirstOrDefault(m => m.Patient.Equals(patientToApply));
                if(mappingForThisPatient != null)
                {
                     _logger.LogDebug("Found existing mapping for patient {PatientName}. Adding selected wells.", patientToApply.Name);
                     mappingForThisPatient.WellPositions = mappingForThisPatient.WellPositions.Union(selectedPositions).ToList();
                     // Ensure the Patient object itself is the updated one if it was an edit
                     if (existingPatient != null && mappingForThisPatient.Patient != existingPatient)
                     {
                        mappingForThisPatient.Patient = existingPatient; // Point back to the edited object
                     }
                }
                else
                {
                     _logger.LogDebug("Creating new mapping for patient {PatientName} with wells: {Wells}", patientToApply.Name, string.Join(",", selectedPositions));
                     PatientMappings.Add(new PatientWellMapping(patientToApply, selectedPositions));
                }
                
                // Update the sample name in the WellInfo objects
                foreach (var well in selectedWells)
                {
                    well.SampleName = patientToApply.Name; // Or MRN or combination
                }
                _logger.LogInformation("Updated sample names for selected wells.");

                 // Ensure commands are cast correctly before notifying
                 (ClearAllPatientsCommand as CommunityToolkit.Mvvm.Input.RelayCommand)?.NotifyCanExecuteChanged();
                 (RemovePatientMappingCommand as CommunityToolkit.Mvvm.Input.RelayCommand<PatientWellMapping>)?.NotifyCanExecuteChanged(); 
            }
            else
            { 
                 _logger.LogInformation("Set patient info cancelled or dialog returned unexpected result.");
                 // Result could be null or false depending on DialogHost setup and Close command parameter
            }
        }
        
        private bool CanRemovePatientMapping(PatientWellMapping? mapping)
        {
            return mapping != null;
        }

        private void RemovePatientMapping(PatientWellMapping? mapping)
        {
            if (!CanRemovePatientMapping(mapping)) return;

            _logger.LogInformation("RemovePatientMappingCommand executed for patient: {PatientName}", mapping!.Patient.Name);
            // Update WellInfo SampleName for the removed wells
            foreach (var wellPos in mapping.WellPositions)
            {
                var well = Wells.FirstOrDefault(w => w.Position == wellPos);
                if (well != null) well.SampleName = string.Empty;
            }
            PatientMappings.Remove(mapping);
            _logger.LogDebug("Patient mapping removed.");

            // Ensure commands are cast correctly before notifying
            (ClearAllPatientsCommand as CommunityToolkit.Mvvm.Input.RelayCommand)?.NotifyCanExecuteChanged();
            (ApplyCommand as CommunityToolkit.Mvvm.Input.RelayCommand)?.NotifyCanExecuteChanged(); // Apply might become possible/impossible
        }

        private bool CanClearAllPatients()
        {
             return PatientMappings.Count > 0;
        }

        private void ClearAllPatients()
        {
            if (!CanClearAllPatients()) return;
            _logger.LogInformation("ClearAllPatientsCommand executed.");

            PatientMappings.Clear();
            // Reset all well SampleNames
            foreach (var well in Wells)
            {
                well.SampleName = string.Empty;
            }
             _logger.LogDebug("All patient mappings and sample names cleared.");
             // Ensure commands are cast correctly before notifying
             (ClearAllPatientsCommand as CommunityToolkit.Mvvm.Input.RelayCommand)?.NotifyCanExecuteChanged();
             (RemovePatientMappingCommand as CommunityToolkit.Mvvm.Input.RelayCommand<PatientWellMapping>)?.NotifyCanExecuteChanged(); // This likely becomes false
             (ApplyCommand as CommunityToolkit.Mvvm.Input.RelayCommand)?.NotifyCanExecuteChanged();
        }
        
        private bool CanApplySampleNames()
        {
            // Enable Apply if there's plate data to apply to
            return _appStateService.CurrentPlate != null; 
        }

        private void ApplySampleNames()
        {
             if (!CanApplySampleNames()) 
             {
                 _logger.LogWarning("ApplyCommand executed but CurrentPlate is null in AppState.");
                 // Show error? 
                 return;
             }

             _logger.LogInformation("ApplyCommand executed. Applying sample names to global plate data.");
            
             var globalPlate = _appStateService.CurrentPlate;
             int updatedCount = 0;
             if (globalPlate != null && globalPlate.WellLayouts != null)
             {
                 foreach (var vmWell in Wells)
                 {
                     var plateWell = globalPlate.WellLayouts.FirstOrDefault(w => w.Position == vmWell.Position);
                     if (plateWell != null)
                     {
                        if (plateWell.SampleName != vmWell.SampleName)
                        {
                            plateWell.SampleName = vmWell.SampleName;
                            // --- Update Associated Patient Info (Example) --- 
                            // Find the mapping for this well in the ViewModel's state
                            var mapping = PatientMappings.FirstOrDefault(m => m.WellPositions.Contains(vmWell.Position));
                            if (mapping != null && mapping.Patient != null)
                            {
                                // Option A: If WellLayout has an AssociatedPatient object (requires cloning to avoid issues)
                                 // plateWell.AssociatedPatient = new PatientInfo(mapping.Patient.Name, mapping.Patient.MedicalRecordNumber);
                                // Option B: If WellLayout has AssociatedPatientId (assuming PatientInfo has Id)
                                // plateWell.AssociatedPatientId = mapping.Patient.Id; 
                                 _logger.LogTrace("Updated patient association for {Position} to Patient Name: {PatientName}", plateWell.Position, mapping.Patient.Name);
                            }
                            else 
                            {
                                // Clear associated patient if sample name was cleared / no mapping found
                                // plateWell.AssociatedPatient = null;
                                // plateWell.AssociatedPatientId = null;
                                 _logger.LogTrace("Cleared patient association for {Position}", plateWell.Position);
                            }
                            // --- End Example --- 
                            updatedCount++;
                        }
                     }
                     else
                     {
                         _logger.LogWarning("Could not find well {Position} in global plate data to apply sample name.", vmWell.Position);
                     }
                 }
                 _logger.LogInformation("Applied sample names to {Count} wells in global plate data.", updatedCount);

                 // --- Navigate to the next step (PCR Result Analysis View) --- 
                 _logger.LogInformation("Navigating to PCR Result Analysis View...");
                 try
                 {
                     _navigationService.NavigateTo<PCRResultAnalysisViewModel>(); // Use the actual ViewModel type
                 }
                 catch(Exception navEx)
                 {
                      _logger.LogError(navEx, "Navigation to PCRResultAnalysisViewModel failed.");
                      // Show error message to user
                 }
                 // --- End Navigation --- 
             }
             else
             {
                  _logger.LogError("ApplyCommand failed: globalPlate or its WellLayouts were null.");
                  // Show error message
             }
        }
    }
} 