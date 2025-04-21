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
                
                // 记录几个孔位的样本，用于调试
                foreach (var well in currentPlate.WellLayouts.Take(10))
                {
                    _logger.LogDebug("Original plate data - Position: {pos}, Row: {row}, Column: {col}, PatientName: {name}", 
                        well.Position, well.Row, well.Column, well.PatientName ?? "null");
                }
                
                // 保留列数设置
                PlateColumns = currentPlate.Columns; 
                
                int defaultRows = 8;
                int defaultColumns = PlateColumns;
                
                // 创建孔板UI视图，但保留之前设置的患者信息
                for (int i = 1; i <= defaultRows; i++) // Rows A-H
                {
                    for (int j = 1; j <= defaultColumns; j++) // Columns 1-12
                    {
                        string rowLetter = ((char)('A' + i - 1)).ToString(); // 行名从A开始
                        string positionString = $"{rowLetter}{j}"; // 例如 "A1"，仅用于查询，不直接赋值
                        
                        // 尝试多种匹配方式查找原始数据
                        var existingWell = currentPlate.WellLayouts.FirstOrDefault(w => 
                            // 1. 直接按Position匹配 (例如 "A1")
                            (w.Position != null && w.Position.Equals(positionString, StringComparison.OrdinalIgnoreCase)) ||
                            // 2. 按Row和Column匹配
                            (w.Row == rowLetter && w.Column == j) ||
                            // 3. 按位置字符串的一部分匹配
                            (w.Position != null && w.Position.Contains(positionString))
                        );
                        
                        if (existingWell != null)
                        {
                            _logger.LogDebug("Found match for {position} - PatientName: {name}", 
                                positionString, existingWell.PatientName ?? "null");
                        }
                        
                        // 创建UI孔位，继承现有的患者信息
                        var well = new WellLayout
                        {
                            Row = rowLetter,
                            Column = j,
                            // Position属性是只读的，不能直接赋值
                            // 它可能是一个计算属性，基于Row和Column自动生成
                            IsSelected = false,
                            // 如果原始数据中有此孔位的患者信息，则使用；否则设为null
                            PatientName = existingWell?.PatientName,
                            PatientCaseNumber = existingWell?.PatientCaseNumber,
                            SampleName = existingWell?.SampleName
                        };
                        
                        // 订阅属性变更事件
                        well.PropertyChanged += Well_PatientInfoChanged;
                        WellLayouts.Add(well);
                    }
                }
                _logger.LogInformation("Created well layout with {Count} wells, preserving existing patient info.", WellLayouts.Count);
                
                // 输出有患者信息的孔位数量，用于调试
                int wellsWithPatient = WellLayouts.Count(w => !string.IsNullOrEmpty(w.PatientName));
                _logger.LogInformation("Found {count} wells with patient information", wellsWithPatient);
                
                // 初始化完成后更新患者列表
                UpdatePatientsList();
            }
            else
            {
                _logger.LogWarning("No plate data found in AppStateService. Generating default 8x12 layout.");
                // Generate a default 8x12 layout if no data exists
                PlateColumns = 12; // Set PlateColumns property
                int defaultRows = 8;
                int defaultColumns = 12;
                for (int i = 1; i <= defaultRows; i++) // Rows A-H
                {
                    for (int j = 1; j <= defaultColumns; j++) // Columns 1-12
                    {
                        var well = new WellLayout
                        {
                             // Make sure WellLayout properties can be set directly or have a suitable constructor
                            Row = ((char)('A' + i - 1)).ToString(), // Convert char to string
                            Column = j, // 列号从1开始
                            // Position是只读属性，不能直接赋值
                            IsSelected = false,
                             // Set other default properties if necessary
                            PatientName = null,
                            PatientCaseNumber = null,
                            SampleName = null,
                             // WellType = CoreWellType.Empty, // Assuming CoreWellType alias exists or use full name
                        };
                         // Subscribe to property changes for the newly created well
                         well.PropertyChanged += Well_PatientInfoChanged;
                         WellLayouts.Add(well);
                    }
                }
                 UpdatePatientsList(); // Update patient list even for empty plate
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

            // 1. 获取必要的分析数据
            var wellData = this.WellLayouts;  // 用户设置的孔位与患者信息
            var currentPlate = _appStateService.CurrentPlate;  // 导入的原始PCR数据
            var analysisMethodPath = _appStateService.CurrentAnalysisMethodPath;  // 选择的分析方法

            if (currentPlate == null || string.IsNullOrEmpty(analysisMethodPath))
            {
                _logger.LogError("无法进行分析：缺少原始数据或分析方法");
                MessageBox.Show("无法进行分析：缺少原始数据或分析方法", "分析错误", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            try
            {
                // 检查是否有孔位未设置患者信息
                var wellsWithoutPatient = wellData.Where(w => string.IsNullOrEmpty(w.PatientName)).ToList();
                if (wellsWithoutPatient.Any())
                {
                    string positions = string.Join(", ", wellsWithoutPatient.Select(w => w.Position));
                    var result = MessageBox.Show(
                        $"以下孔位未设置患者信息：{positions}\n是否继续分析？", 
                        "缺少患者信息", 
                        MessageBoxButton.YesNo, 
                        MessageBoxImage.Warning);
                    
                    if (result == MessageBoxResult.No)
                    {
                        return;
                    }
                }

                // 2. 更新当前孔位信息到全局数据
                // 这里我们需要将用户设置的患者信息更新到原始数据中
                foreach (var uiWell in wellData)
                {
                    // 查找对应的原始数据孔位
                    var originalWell = currentPlate.WellLayouts.FirstOrDefault(w => 
                        w.Row == uiWell.Row && w.Column == uiWell.Column);
                    
                    if (originalWell != null)
                    {
                        // 更新患者信息
                        originalWell.PatientName = uiWell.PatientName;
                        originalWell.PatientCaseNumber = uiWell.PatientCaseNumber;
                        originalWell.SampleName = uiWell.SampleName;
                    }
                }

                _logger.LogInformation("已将用户设置的患者信息更新到原始数据中");

                // 3. 导航到结果分析视图
                _logger.LogInformation("Navigating to PCR Result Analysis View...");
                _navigationService.NavigateTo<PCRResultAnalysisViewModel>(); 
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "分析过程中发生错误");
                MessageBox.Show("分析过程中发生错误: " + ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // TODO: Implement logic for drag selection and keyboard shortcuts
        // These often involve interaction with the View's code-behind or attached behaviors.
    }
} 