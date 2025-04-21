using MolePCRConvert4WPF.App.Commands;
using MolePCRConvert4WPF.Core.Interfaces;
using MolePCRConvert4WPF.Core.Models;
using CommunityToolkit.Mvvm.Input; // Added for IAsyncRelayCommand
using System.Collections.ObjectModel;
using System.Windows.Input;
using Microsoft.Win32; // For OpenFileDialog
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using System; // For ArgumentNullException, Enum
using System.Linq; // For Enum.GetValues, Where, FirstOrDefault
using System.Windows; // For MessageBox
using System.Collections.Generic; // For List
using MolePCRConvert4WPF.Core.Enums; // For InstrumentType
using MolePCRConvert4WPF.Core.Services; // Assuming INavigationService is here
using MolePCRConvert4WPF.Infrastructure.Services; // Use this for NavigationService impl
using MolePCRConvert4WPF.Core.Models; // <-- Add using for FileDisplayInfo -->
using System.IO; // For Directory, Path

namespace MolePCRConvert4WPF.App.ViewModels
{
    public class DataInputViewModel : ViewModelBase
    {
        private readonly IFileHandler _fileHandler;
        private readonly ILogger<DataInputViewModel> _logger;
        private readonly IAppStateService _appStateService; // <-- Inject AppStateService -->
        private readonly IUserSettingsService _userSettingsService; // <-- Add field for UserSettingsService -->
        private string _selectedFilePath = string.Empty; // Initialize to empty string
        private bool _isLoading;
        private ObservableCollection<InstrumentType> _instrumentTypes;
        private InstrumentType _selectedInstrumentType;
        private ObservableCollection<string> _availableAnalysisMethods; // Added
        private string? _selectedAnalysisMethod; // Added
        private readonly INavigationService _navigationService;
        
        public string SelectedFilePath
        {
            get => _selectedFilePath;
            // Update CanExecute when path changes
            set 
            { 
                if (SetProperty(ref _selectedFilePath, value))
                {
                    // Inform the command that its CanExecute status might have changed
                    LoadDataCommand?.NotifyCanExecuteChanged(); 
                    // Automatically detect type when path changes
                    DetectInstrumentType();
                }
            }
        }

        public bool IsLoading
        {
            get => _isLoading;
            // Update CanExecute when loading state changes
            set 
            {
                if (SetProperty(ref _isLoading, value))
                {
                    LoadDataCommand?.NotifyCanExecuteChanged();
                    (BrowseCommand as CommunityToolkit.Mvvm.Input.RelayCommand)?.NotifyCanExecuteChanged();
                    (CancelCommand as CommunityToolkit.Mvvm.Input.RelayCommand)?.NotifyCanExecuteChanged();
                }
            }
        }
        
        public ObservableCollection<InstrumentType> InstrumentTypes
        {
            get => _instrumentTypes;
            // Private set might be better if only initialized in constructor
            private set => SetProperty(ref _instrumentTypes, value);
        }

        public InstrumentType SelectedInstrumentType
        {
            get => _selectedInstrumentType;
            set => SetProperty(ref _selectedInstrumentType, value);
        }
        
        public ObservableCollection<string> AvailableAnalysisMethods
        {
            get => _availableAnalysisMethods;
            private set => SetProperty(ref _availableAnalysisMethods, value);
        }

        public string? SelectedAnalysisMethod
        {
            get => _selectedAnalysisMethod;
            // Update CanExecute when method changes
            set { 
                if (SetProperty(ref _selectedAnalysisMethod, value)) 
                { 
                    LoadDataCommand?.NotifyCanExecuteChanged(); 
                } 
            }
        }
        
        public ICommand BrowseCommand { get; }
        public IAsyncRelayCommand LoadDataCommand { get; }
        public ICommand CancelCommand { get; }

        public DataInputViewModel(
            IFileHandler fileHandler, 
            ILogger<DataInputViewModel> logger, 
            IAppStateService appStateService, // <-- Add IAppStateService -->
            INavigationService navigationService,
            IUserSettingsService userSettingsService) // <-- Add IUserSettingsService -->
        {
            _fileHandler = fileHandler ?? throw new ArgumentNullException(nameof(fileHandler));
            _logger = logger;
            _appStateService = appStateService; // <-- Assign injected service -->
            _navigationService = navigationService;
            _userSettingsService = userSettingsService; // <-- Assign injected service -->
            // Initialize InstrumentTypes here
            _instrumentTypes = new ObservableCollection<InstrumentType>(Enum.GetValues<InstrumentType>().Where(t => t != InstrumentType.Unknown));
            SelectedInstrumentType = _instrumentTypes.FirstOrDefault();
            // _selectedFilePath is initialized at declaration
            
            // Initialize AvailableAnalysisMethods
            _availableAnalysisMethods = new ObservableCollection<string>(); // Initialize empty
            
            // 优先从用户设置中加载分析方法文件夹路径，并读取该文件夹中的方法文件
            LoadAnalysisMethodsFromSettings();
            
            BrowseCommand = new CommunityToolkit.Mvvm.Input.RelayCommand(BrowseFile, () => !IsLoading);
            LoadDataCommand = new AsyncRelayCommand(LoadDataAsync, CanLoadData);
            CancelCommand = new CommunityToolkit.Mvvm.Input.RelayCommand(Cancel, CanCancel);

            // Consider subscribing to AppState changes if needed for real-time updates
            // For now, updating on init and potentially on activation/re-navigation is simpler
        }

        /// <summary>
        /// 从用户设置加载分析方法文件夹路径，并读取该文件夹中的分析方法文件
        /// </summary>
        private void LoadAnalysisMethodsFromSettings()
        {
            _logger?.LogInformation("尝试从用户设置加载分析方法文件夹路径");
            string? savedFolderPath = _userSettingsService.LoadSetting("AnalysisMethodFolderPath");
            
            if (!string.IsNullOrEmpty(savedFolderPath) && Directory.Exists(savedFolderPath))
            {
                _logger?.LogInformation($"从用户设置加载分析方法文件夹路径: {savedFolderPath}");
                try
                {
                    // 从文件夹加载方法文件
                    var files = Directory.GetFiles(savedFolderPath, "*.xlsx")
                                         .Select(f => new FileDisplayInfo { DisplayName = Path.GetFileName(f), FullPath = f })
                                         .OrderBy(f => f.DisplayName)
                                         .ToList();
                    
                    // 更新AppStateService中的方法文件列表
                    if (_appStateService.AnalysisMethodFiles == null)
                    {
                        _appStateService.AnalysisMethodFiles = new ObservableCollection<FileDisplayInfo>();
                    }
                    else
                    {
                        _appStateService.AnalysisMethodFiles.Clear();
                    }
                    
                    foreach (var file in files)
                    {
                        _appStateService.AnalysisMethodFiles.Add(file);
                    }
                    
                    _logger?.LogInformation($"成功从文件夹加载了 {files.Count} 个分析方法文件");
                    
                    // 更新可用方法下拉列表
                    UpdateAvailableMethods();
                }
                catch (Exception ex)
                {
                    _logger?.LogError(ex, $"从文件夹 {savedFolderPath} 加载分析方法文件时出错");
                }
            }
            else
            {
                _logger?.LogInformation("未找到有效的分析方法文件夹路径设置，将使用AppStateService中的方法列表");
                // 使用AppStateService中的方法列表
                UpdateAvailableMethods();
            }
        }

        private void UpdateAvailableMethods()
        {
            _logger.LogDebug("Updating AvailableAnalysisMethods from AppStateService...");
            var currentSelection = SelectedAnalysisMethod; // Preserve selection if possible
            
            AvailableAnalysisMethods.Clear();
            if (_appStateService.AnalysisMethodFiles != null && _appStateService.AnalysisMethodFiles.Any())
            {
                foreach (var fileInfo in _appStateService.AnalysisMethodFiles)
                {
                    AvailableAnalysisMethods.Add(fileInfo.DisplayName);
                }
                _logger.LogDebug("Found {Count} methods in AppState.", AvailableAnalysisMethods.Count);
            }
            else
            {
                 _logger.LogDebug("No analysis method files found in AppState.");
                 // Optionally add a default or placeholder
                 // AvailableAnalysisMethods.Add("无可用方法");
            }

            // Try to restore selection or select the first
            SelectedAnalysisMethod = AvailableAnalysisMethods.Contains(currentSelection) 
                                        ? currentSelection 
                                        : AvailableAnalysisMethods.FirstOrDefault();
            _logger.LogDebug("Selected Analysis Method after update: {Method}", SelectedAnalysisMethod ?? "None");
        }

        // Call UpdateAvailableMethods when the view becomes active 
        // (This requires lifecycle events or a messaging system, 
        // for now, it's updated only on construction)
        public void OnNavigatedTo() // Placeholder for activation logic
        {
             // 每次导航到该视图时都尝试从用户设置加载分析方法文件
             LoadAnalysisMethodsFromSettings();
        }

        private void BrowseFile()
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls|All files (*.*)|*.*",
                Title = "选择PCR数据文件"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                SelectedFilePath = openFileDialog.FileName; // This setter will trigger CanExecuteChanged and DetectInstrumentType
            }
        }
        
        private async void DetectInstrumentType()
        {
             if (string.IsNullOrEmpty(SelectedFilePath)) return;
             try
             {
                 var detectedType = await _fileHandler.DetectInstrumentTypeAsync(SelectedFilePath);
                 if (InstrumentTypes.Contains(detectedType))
                 {
                     SelectedInstrumentType = detectedType;
                 }
                 else
                 {
                     _logger?.LogWarning($"检测到的仪器类型 {detectedType} 不在支持列表中，将使用默认类型。");
                     // Keep the default or first type selected
                 }
             }
             catch(Exception ex)
             {
                 _logger?.LogError(ex, "自动检测仪器类型时出错");
                 // Handle error, maybe show message to user
             }
        }

        private bool CanLoadData()
        {
            return !string.IsNullOrEmpty(SelectedFilePath) && 
                   SelectedInstrumentType != InstrumentType.Unknown &&
                   !string.IsNullOrEmpty(SelectedAnalysisMethod) &&
                   !IsLoading;
        }

        private async Task LoadDataAsync()
        {
            if (!CanLoadData()) return;

            IsLoading = true;
            LoadDataCommand?.NotifyCanExecuteChanged();
            (BrowseCommand as CommunityToolkit.Mvvm.Input.RelayCommand)?.NotifyCanExecuteChanged();
            (CancelCommand as CommunityToolkit.Mvvm.Input.RelayCommand)?.NotifyCanExecuteChanged();
            _logger?.LogInformation("LoadDataCommand started. File: {FilePath}, Instrument: {InstrumentType}, Method: {AnalysisMethod}", SelectedFilePath, SelectedInstrumentType, SelectedAnalysisMethod);

            try
            {
                // Simulate data loading and processing
                 _logger?.LogDebug("Calling FileHandler.ReadPcrDataAsync...");
                 // Pass Guid.NewGuid() for plateId - adjust if specific ID logic is needed
                 var plateId = Guid.NewGuid(); 
                 _logger.LogDebug("Generated Plate ID: {PlateId}", plateId);
                 var loadedPlate = await _fileHandler.ReadPcrDataAsync(SelectedFilePath, SelectedInstrumentType, plateId);
                 _logger?.LogDebug("FileHandler.ReadPcrDataAsync completed.");

                // --- Check if plate loading was successful --- 
                if (loadedPlate != null)
                {
                    // Store plate and *selected method path* in AppState
                    _appStateService.CurrentPlate = loadedPlate;
                    // Find the full path corresponding to the selected display name
                    string? selectedMethodFullPath = _appStateService.AnalysisMethodFiles?
                                                       .FirstOrDefault(f => f.DisplayName == SelectedAnalysisMethod)?.FullPath;
                                                       
                    // Check if a valid method path was found
                    if(string.IsNullOrEmpty(selectedMethodFullPath))
                    {
                        _logger?.LogError("未能找到与所选分析方法名称 '{SelectedMethodName}' 匹配的完整文件路径。", SelectedAnalysisMethod);
                        MessageBox.Show($"错误：无法找到分析方法 '{SelectedAnalysisMethod}' 的文件路径。请检查方法列表或配置。", "分析方法错误", MessageBoxButton.OK, MessageBoxImage.Warning);
                        // Optionally clear the selected method or handle differently
                         _appStateService.CurrentAnalysisMethodPath = null; // Ensure it's null if not found
                         // Do not navigate if method path is invalid
                         return; 
                    }
                    
                    _appStateService.CurrentAnalysisMethodPath = selectedMethodFullPath;
                    
                    _logger?.LogInformation("Data loaded successfully. Plate ID: {PlateId}. Stored Method Path: {MethodPath}", 
                                            loadedPlate.Id, _appStateService.CurrentAnalysisMethodPath ?? "Not Set");
                    
                    // --- Navigate to Sample Naming/Analysis View ---
                    _logger?.LogInformation("Navigating to Sample Analysis View...");
                    _navigationService.NavigateTo<SampleAnalysisViewModel>(); 
                }
                else
                {
                    // Handle the case where ReadPcrDataAsync returned null (due to error)
                    _logger?.LogError("LoadDataAsync failed because FileHandler returned null plate data. File: {FilePath}", SelectedFilePath);
                    // Show error message to user 
                    MessageBox.Show($"加载文件 '{System.IO.Path.GetFileName(SelectedFilePath)}' 时出错。\n请检查文件格式、内容以及选择的仪器类型是否正确。\n错误详情请查看日志。", "数据加载失败", MessageBoxButton.OK, MessageBoxImage.Error);
                    // Do NOT navigate
                }
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "Error during data loading process.");
                // Show error message to user
                 // _dialogService.ShowError($"An error occurred: {ex.Message}");
            }
            finally
            {
                IsLoading = false;
                LoadDataCommand?.NotifyCanExecuteChanged();
                (BrowseCommand as CommunityToolkit.Mvvm.Input.RelayCommand)?.NotifyCanExecuteChanged();
                (CancelCommand as CommunityToolkit.Mvvm.Input.RelayCommand)?.NotifyCanExecuteChanged();
            }
        }

        private bool CanCancel()
        {
             // Example: Allow cancel only if NOT loading?
             // return !IsLoading;
             return true; // Or implement specific cancel logic
        }
        
        private void Cancel()
        {
            _logger?.LogInformation("CancelCommand executed.");
            // Implement cancel logic - e.g., close the view/dialog, navigate back
            // Example: _navigationService.GoBack();
            // Example: RequestClose?.Invoke(); // If it's a dialog VM
        }
    }
} 