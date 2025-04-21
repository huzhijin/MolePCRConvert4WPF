using MolePCRConvert4WPF.App.Commands;
using MolePCRConvert4WPF.Core.Interfaces;
using MolePCRConvert4WPF.Core.Models;
using System.Collections.ObjectModel;
using System.Windows.Input;
using System.IO;
using System.Linq;
using Microsoft.Extensions.Logging;
using System;
using System.Threading.Tasks;
using System.Collections.Generic; // For List
using Ookii.Dialogs.Wpf; // Use Ookii for Folder Browser Dialog
using System.Windows; // For MessageBox
using MolePCRConvert4WPF.Core.Services; // <-- Add using for IAppStateService

namespace MolePCRConvert4WPF.App.ViewModels
{
    public class ExcelAnalysisMethodConfigViewModel : ViewModelBase
    {
        private readonly ILogger<ExcelAnalysisMethodConfigViewModel> _logger;
        private readonly IAppStateService _appStateService; // <-- Add field for AppStateService
        private readonly IAnalysisMethodConfigService _configService; // <-- Add field for ConfigService -->
        private readonly IUserSettingsService _userSettingsService; // <-- Add field for UserSettingsService -->

        private string? _methodFolderPath;
        private ObservableCollection<FileDisplayInfo> _excelFiles = new ObservableCollection<FileDisplayInfo>();
        private FileDisplayInfo? _selectedExcelFile;
        private string? _currentFileName;
        private ObservableCollection<AnalysisMethodRule> _methodConfiguration = new ObservableCollection<AnalysisMethodRule>();
        private bool _isLoading;

        public string? MethodFolderPath
        {
            get => _methodFolderPath;
            set
            {
                string? directoryPath = value;
                if (!string.IsNullOrEmpty(directoryPath))
                {
                    // Check if the path points to a file, if so, get the directory
                    if (File.Exists(directoryPath))
                    {
                        directoryPath = Path.GetDirectoryName(directoryPath);
                    }
                    // Check if it might be a directory path already ending with separator
                    else if (!Directory.Exists(directoryPath) && (directoryPath.EndsWith(Path.DirectorySeparatorChar.ToString()) || directoryPath.EndsWith(Path.AltDirectorySeparatorChar.ToString())))
                    {
                         // Possibly an incomplete path, try removing trailing separator before check
                        var tempPath = directoryPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
                        if (!Directory.Exists(tempPath))
                        {
                             // If it's still not a valid directory, maybe clear it or log warning
                             // For now, we'll proceed hoping it's intended as a directory path
                        }
                    }
                     // If it's not an existing file or directory, assume it's intended as a directory path.
                }
                
                if (SetProperty(ref _methodFolderPath, directoryPath))
                {
                    // Update AppState with the cleaned directory path
                    if (_appStateService != null) // Ensure service is available
                    { 
                       _appStateService.CurrentAnalysisMethodPath = _methodFolderPath;
                    }
                    // Refresh files when path changes
                     RefreshExcelFiles();
                     // Update dependent command states
                     (RefreshFilesCommand as RelayCommand)?.RaiseCanExecuteChanged();
                     (CreateNewFileCommand as RelayCommand)?.RaiseCanExecuteChanged();
                }
            }
        }

        public ObservableCollection<FileDisplayInfo> ExcelFiles
        {
            get => _excelFiles;
            private set => SetProperty(ref _excelFiles, value);
        }

        public FileDisplayInfo? SelectedExcelFile
        {
            get => _selectedExcelFile;
            set
            {
                if (SetProperty(ref _selectedExcelFile, value))
                {
                    CurrentFileName = value?.DisplayName;
                    // Load the configuration for the selected file
                    LoadConfigurationAsync(value?.FullPath);
                }
            }
        }

        public string? CurrentFileName
        {
            get => _currentFileName;
            private set => SetProperty(ref _currentFileName, value);
        }

        public ObservableCollection<AnalysisMethodRule> MethodConfiguration
        {
            get => _methodConfiguration;
            private set => SetProperty(ref _methodConfiguration, value);
        }

        public bool IsLoading
        {
            get => _isLoading;
            private set
            {
                if(SetProperty(ref _isLoading, value))
                {
                    // Update command states when loading changes
                    (BrowseFolderCommand as RelayCommand)?.RaiseCanExecuteChanged();
                    (RefreshFilesCommand as RelayCommand)?.RaiseCanExecuteChanged();
                    (CreateNewFileCommand as RelayCommand)?.RaiseCanExecuteChanged();
                    (SaveConfigurationCommand as RelayCommand)?.RaiseCanExecuteChanged();
                }
            }
        }

        public ICommand BrowseFolderCommand { get; }
        public ICommand RefreshFilesCommand { get; }
        public ICommand CreateNewFileCommand { get; }
        public ICommand SaveConfigurationCommand { get; }
        // Close command might be handled by the View or Window itself

        public ExcelAnalysisMethodConfigViewModel(ILogger<ExcelAnalysisMethodConfigViewModel> logger,
                                                IAppStateService appStateService, 
                                                IAnalysisMethodConfigService configService,
                                                IUserSettingsService userSettingsService) // <-- Add parameter -->
        {
            _logger = logger;
            _appStateService = appStateService; 
            _configService = configService; // <-- Assign injected service -->
            _userSettingsService = userSettingsService; // <-- Assign injected service -->

            BrowseFolderCommand = new RelayCommand(BrowseFolder, () => !IsLoading);
            RefreshFilesCommand = new RelayCommand(RefreshExcelFiles, () => !string.IsNullOrEmpty(MethodFolderPath) && !IsLoading);
            CreateNewFileCommand = new RelayCommand(CreateNewExcel, () => !string.IsNullOrEmpty(MethodFolderPath) && !IsLoading);
            SaveConfigurationCommand = new RelayCommand(async () => await SaveConfigurationAsync(), () => SelectedExcelFile != null && !IsLoading);

            // 优先从用户设置中加载上次选择的文件夹路径
            string? savedFolderPath = _userSettingsService.LoadSetting("AnalysisMethodFolderPath");
            
            if (!string.IsNullOrEmpty(savedFolderPath) && Directory.Exists(savedFolderPath))
            {
                // 如果有保存的路径且存在，则使用该路径
                MethodFolderPath = savedFolderPath;
                _logger?.LogInformation($"从用户设置加载分析方法文件夹路径: {savedFolderPath}");
            }
            else
            {
                // 否则尝试从AppState加载
                string? initialPath = _appStateService.CurrentAnalysisMethodPath;
                if (!string.IsNullOrEmpty(initialPath) && File.Exists(initialPath)) // Check if it's a file path
                {
                    MethodFolderPath = Path.GetDirectoryName(initialPath); // Use only the directory
                }
                else
                {
                    MethodFolderPath = initialPath; // Assume it's already a directory path or null/empty
                }
            }
            // Note: RefreshExcelFiles() is called within the MethodFolderPath setter now
        }

        private void BrowseFolder()
        {
            var dialog = new VistaFolderBrowserDialog
            {
                Description = "选择分析方法文件所在的文件夹",
                UseDescriptionForTitle = true,
                SelectedPath = MethodFolderPath ?? Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            };

            if (dialog.ShowDialog() == true)
            {
                // The setter for MethodFolderPath now handles storing to AppState and refreshing
                MethodFolderPath = dialog.SelectedPath;
                
                // 保存选择的文件夹路径到用户设置
                if (!string.IsNullOrEmpty(MethodFolderPath) && Directory.Exists(MethodFolderPath))
                {
                    _userSettingsService.SaveSetting("AnalysisMethodFolderPath", MethodFolderPath);
                    _logger?.LogInformation($"已保存分析方法文件夹路径到用户设置: {MethodFolderPath}");
                }
            }
        }

        private void RefreshExcelFiles()
        {
            // Clear existing files in AppState first
            _appStateService.AnalysisMethodFiles?.Clear();
            
            if (string.IsNullOrEmpty(MethodFolderPath) || !Directory.Exists(MethodFolderPath))
            {
                ExcelFiles.Clear();
                SelectedExcelFile = null;
                return;
            }

            _logger?.LogInformation($"Refreshing files in: {MethodFolderPath}");
            try
            {
                var files = Directory.GetFiles(MethodFolderPath, "*.xlsx")
                                     .Select(f => new FileDisplayInfo { DisplayName = Path.GetFileName(f), FullPath = f })
                                     .OrderBy(f => f.DisplayName)
                                     .ToList();
                                     
                ExcelFiles = new ObservableCollection<FileDisplayInfo>(files);
                // <-- Update AppStateService with the new list -->
                _appStateService.AnalysisMethodFiles = ExcelFiles; 
                
                // Optionally re-select the previously selected file if it still exists
                var previouslySelectedPath = SelectedExcelFile?.FullPath;
                SelectedExcelFile = ExcelFiles.FirstOrDefault(f => f.FullPath == previouslySelectedPath) ?? ExcelFiles.FirstOrDefault();

            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, $"刷新文件夹时出错: {MethodFolderPath}");
                MessageBox.Show($"无法读取文件夹内容: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                ExcelFiles.Clear();
                SelectedExcelFile = null;
                // Also clear AppState on error
                _appStateService.AnalysisMethodFiles?.Clear();
            }
        }

        private void CreateNewExcel()
        {
            // TODO: Implement logic to create a new, empty Excel file (or based on a template)
            // Requires user input for the filename
            // After creation, refresh the list and select the new file
            _logger?.LogWarning("创建新文件功能尚未实现。");
            MessageBox.Show("创建新分析方法文件的功能正在开发中。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private async Task LoadConfigurationAsync(string? filePath)
        {
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                MethodConfiguration.Clear();
                return;
            }

            IsLoading = true;
            MethodConfiguration.Clear(); // Clear previous data before loading
            _logger?.LogInformation($"正在加载配置文件: {filePath}");
            try
            {
                // <-- Use the injected service to load data -->
                var config = await _configService.LoadConfigurationAsync(filePath);
                MethodConfiguration = new ObservableCollection<AnalysisMethodRule>(config);
                _logger?.LogInformation("配置文件加载成功，共 {Count} 条规则。", MethodConfiguration.Count);
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, $"加载配置文件时出错: {filePath}");
                MessageBox.Show($"无法加载配置文件: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                MethodConfiguration.Clear(); // Ensure clear on error
            }
            finally
            {
                IsLoading = false;
                // Ensure Save command CanExecute updates after loading attempt
                (SaveConfigurationCommand as RelayCommand)?.RaiseCanExecuteChanged(); 
            }
        }

        private async Task SaveConfigurationAsync()
        {
            // TODO: Implement logic to save the configuration to a file
            _logger?.LogWarning("保存配置文件功能尚未实现。");
            MessageBox.Show("保存分析方法配置的功能正在开发中。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }
} 