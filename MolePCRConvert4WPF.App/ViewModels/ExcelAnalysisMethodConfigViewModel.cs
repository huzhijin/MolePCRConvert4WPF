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
using System.Windows.Controls;
using System.Windows.Media;

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
        private DataGrid? _dataGrid;

        // 配置文件夹路径
        private const string CONFIG_FOLDER = "Config";

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
                    (RefreshFilesCommand as RelayCommand)?.RaiseCanExecuteChanged();
                    (CreateNewFileCommand as RelayCommand)?.RaiseCanExecuteChanged();
                    (SaveConfigurationCommand as RelayCommand)?.RaiseCanExecuteChanged();
                }
            }
        }

        public ICommand RefreshFilesCommand { get; }
        public ICommand CreateNewFileCommand { get; }
        public ICommand SaveConfigurationCommand { get; }
        public ICommand ImportFileCommand { get; }
        public ICommand ExportFileCommand { get; }
        public ICommand AddNewRowCommand { get; }
        public ICommand DeleteSelectedRowsCommand { get; }
        public ICommand DeleteFileCommand { get; }
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

            RefreshFilesCommand = new RelayCommand(RefreshExcelFiles, () => !IsLoading);
            CreateNewFileCommand = new RelayCommand(CreateNewExcel, () => !IsLoading);
            SaveConfigurationCommand = new RelayCommand(async () => await SaveConfigurationAsync(), () => SelectedExcelFile != null && !IsLoading);
            ImportFileCommand = new RelayCommand(ImportFile, () => !IsLoading);
            ExportFileCommand = new RelayCommand(ExportFile, () => SelectedExcelFile != null && !IsLoading);
            AddNewRowCommand = new RelayCommand(AddNewRow, () => SelectedExcelFile != null && !IsLoading);
            DeleteSelectedRowsCommand = new RelayCommand(DeleteSelectedRows, () => SelectedExcelFile != null && !IsLoading);
            DeleteFileCommand = new RelayCommand(DeleteFile, () => SelectedExcelFile != null && !IsLoading);

            // 初始化Config文件夹路径
            InitializeConfigFolder();
            
            // 加载Config文件夹中的文件
            MethodFolderPath = GetConfigFolderPath();
        }

        // 获取Config文件夹路径
        private string GetConfigFolderPath()
        {
            // 获取应用程序所在目录
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            // 构建Config文件夹的完整路径
            string configPath = Path.Combine(baseDir, CONFIG_FOLDER);
            return configPath;
        }

        // 初始化Config文件夹
        private void InitializeConfigFolder()
        {
            string configPath = GetConfigFolderPath();
            
            // 确保Config文件夹存在
            if (!Directory.Exists(configPath))
            {
                try
                {
                    Directory.CreateDirectory(configPath);
                    _logger?.LogInformation($"创建Config文件夹: {configPath}");
                }
                catch (Exception ex)
                {
                    _logger?.LogError(ex, $"创建Config文件夹失败: {configPath}");
                    MessageBox.Show($"无法创建配置文件夹: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }

            // 将路径保存到用户设置中
            _userSettingsService.SaveSetting("AnalysisMethodFolderPath", configPath);
            _logger?.LogInformation($"已保存分析方法文件夹路径到用户设置: {configPath}");
        }

        private void RefreshExcelFiles()
        {
            // 清空AppState中的现有文件
            _appStateService.AnalysisMethodFiles?.Clear();
            
            string configPath = GetConfigFolderPath();
            if (!Directory.Exists(configPath))
            {
                // 如果Config文件夹不存在，尝试创建它
                try
                {
                    Directory.CreateDirectory(configPath);
                    _logger?.LogInformation($"刷新时创建Config文件夹: {configPath}");
                }
                catch (Exception ex)
                {
                    _logger?.LogError(ex, $"刷新时创建Config文件夹失败: {configPath}");
                    MessageBox.Show($"无法创建配置文件夹: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    ExcelFiles.Clear();
                    SelectedExcelFile = null;
                    return;
                }
            }

            // 更新MethodFolderPath，确保使用Config文件夹路径
            if (_methodFolderPath != configPath)
            {
                _methodFolderPath = configPath;
                _appStateService.CurrentAnalysisMethodPath = _methodFolderPath;
            }

            _logger?.LogInformation($"正在刷新文件夹中的文件: {_methodFolderPath}");
            
            // 记住当前选中的文件路径
            string? previousSelectedPath = SelectedExcelFile?.FullPath;
            string? previousSelectedFileName = SelectedExcelFile?.DisplayName;
            
            try
            {
                // 获取文件夹中的所有Excel文件
                var allFiles = Directory.GetFiles(_methodFolderPath, "*.xlsx")
                                      .Select(f => new FileDisplayInfo { DisplayName = Path.GetFileName(f), FullPath = f });
                
                // 过滤掉临时文件 
                var validFiles = allFiles.Where(f => !f.DisplayName.StartsWith("temp_"))
                                        .OrderBy(f => f.DisplayName)
                                        .ToList();
                
                // 更新文件列表                
                ExcelFiles = new ObservableCollection<FileDisplayInfo>(validFiles);
                
                // 更新AppStateService中的方法文件列表
                _appStateService.AnalysisMethodFiles = ExcelFiles;
                
                // 尝试重新选择之前选中的文件
                FileDisplayInfo? fileToSelect = null;
                
                // 首先按完整路径匹配
                if (!string.IsNullOrEmpty(previousSelectedPath))
                {
                    fileToSelect = ExcelFiles.FirstOrDefault(f => 
                        f.FullPath.Equals(previousSelectedPath, StringComparison.OrdinalIgnoreCase));
                }
                
                // 如果没找到，则按文件名匹配
                if (fileToSelect == null && !string.IsNullOrEmpty(previousSelectedFileName))
                {
                    // 去掉可能的临时文件前缀
                    string cleanFileName = previousSelectedFileName;
                    if (cleanFileName.StartsWith("temp_"))
                    {
                        cleanFileName = cleanFileName.Substring(5);
                    }
                    
                    fileToSelect = ExcelFiles.FirstOrDefault(f => 
                        f.DisplayName.Equals(cleanFileName, StringComparison.OrdinalIgnoreCase));
                }
                
                // 如果还是没找到，则选择第一个文件
                SelectedExcelFile = fileToSelect ?? ExcelFiles.FirstOrDefault();
                
                _logger?.LogInformation($"已刷新文件夹，共 {ExcelFiles.Count} 个文件");
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, $"刷新文件夹时出错: {_methodFolderPath}");
                MessageBox.Show($"无法读取文件夹内容: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                ExcelFiles.Clear();
                SelectedExcelFile = null;
                
                // 出错时也清空AppState
                _appStateService.AnalysisMethodFiles?.Clear();
            }
        }

        private void CreateNewExcel()
        {
            string configPath = GetConfigFolderPath();
            if (!Directory.Exists(configPath))
            {
                try
                {
                    Directory.CreateDirectory(configPath);
                    _logger?.LogInformation($"创建新文件时创建Config文件夹: {configPath}");
                }
                catch (Exception ex)
                {
                    _logger?.LogError(ex, $"创建新文件时创建Config文件夹失败: {configPath}");
                    MessageBox.Show($"无法创建配置文件夹: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
            }

            // 更新MethodFolderPath，确保使用Config文件夹路径
            if (_methodFolderPath != configPath)
            {
                _methodFolderPath = configPath;
                _appStateService.CurrentAnalysisMethodPath = _methodFolderPath;
            }

            // 创建自定义输入对话框
            var inputDialog = new Window
            {
                Title = "创建新的分析方法文件",
                Width = 400,
                Height = 180,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                ResizeMode = ResizeMode.NoResize,
                WindowStyle = WindowStyle.ToolWindow,
                Owner = Application.Current.MainWindow
            };

            var grid = new Grid { Margin = new Thickness(10) };
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            var label = new TextBlock
            {
                Text = "请输入新文件名称:",
                Margin = new Thickness(0, 0, 0, 10),
                FontSize = 14
            };
            Grid.SetRow(label, 0);

            var textBox = new TextBox
            {
                Margin = new Thickness(0, 0, 0, 20),
                FontSize = 14,
                Text = "新分析方法_" + DateTime.Now.ToString("yyyyMMdd"),
                HorizontalAlignment = HorizontalAlignment.Stretch
            };
            Grid.SetRow(textBox, 1);

            var buttonPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right
            };
            Grid.SetRow(buttonPanel, 2);

            var okButton = new Button
            {
                Content = "确定",
                IsDefault = true,
                Width = 80,
                Height = 30,
                Margin = new Thickness(0, 0, 10, 0)
            };

            var cancelButton = new Button
            {
                Content = "取消",
                IsCancel = true,
                Width = 80,
                Height = 30
            };

            buttonPanel.Children.Add(okButton);
            buttonPanel.Children.Add(cancelButton);

            grid.Children.Add(label);
            grid.Children.Add(textBox);
            grid.Children.Add(buttonPanel);

            inputDialog.Content = grid;

            bool? result = null;
            okButton.Click += (s, e) =>
            {
                result = true;
                inputDialog.Close();
            };

            inputDialog.ShowDialog();

            if (result == true)
            {
                string fileName = textBox.Text.Trim();
                
                // 确保文件名有后缀
                if (!fileName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                {
                    fileName += ".xlsx";
                }

                // 完整文件路径
                string filePath = Path.Combine(_methodFolderPath, fileName);

                IsLoading = true;
                try
                {
                    // 使用配置服务创建新的Excel文件
                    _configService.CreateNewConfigurationFileAsync(filePath).ContinueWith(task =>
                    {
                        // 返回UI线程
                        Application.Current.Dispatcher.Invoke(() =>
                        {
                            IsLoading = false;
                            
                            if (task.IsCompletedSuccessfully)
                            {
                                _logger?.LogInformation($"成功创建新的分析方法文件: {filePath}");
                                // 刷新文件列表
                                RefreshExcelFiles();
                                // 选择新创建的文件
                                SelectedExcelFile = ExcelFiles.FirstOrDefault(f => f.FullPath == filePath);
                                
                                // 创建默认配置项以便用户直接编辑
                                if (MethodConfiguration.Count == 0)
                                {
                                    var defaultRule = new AnalysisMethodRule
                                    {
                                        Index = 1,
                                        WellPosition = "A1",
                                        Channel = "FAM",
                                        SpeciesName = "请输入种名",
                                        JudgeFormula = "[CT]<36",
                                        ConcentrationFormula = ""
                                    };
                                    MethodConfiguration.Add(defaultRule);
                                }
                            }
                            else if (task.Exception != null)
                            {
                                _logger?.LogError(task.Exception, $"创建新文件时出错: {filePath}");
                                MessageBox.Show($"创建文件失败: {task.Exception.InnerException?.Message ?? task.Exception.Message}", 
                                               "创建错误", MessageBoxButton.OK, MessageBoxImage.Error);
                            }
                        });
                    });
                }
                catch (Exception ex)
                {
                    IsLoading = false;
                    _logger?.LogError(ex, $"创建新文件过程中出错: {filePath}");
                    MessageBox.Show($"创建文件过程中出错: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private async Task LoadConfigurationAsync(string? filePath)
        {
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                MethodConfiguration.Clear();
                return;
            }

            // 检查文件名是否是临时文件，如果是则尝试找到原始文件
            string fileName = Path.GetFileName(filePath);
            if (fileName.StartsWith("temp_") && MethodFolderPath != null)
            {
                string originalFileName = fileName.Substring(5); // 去掉"temp_"前缀
                string potentialOriginalPath = Path.Combine(MethodFolderPath, originalFileName);
                
                if (File.Exists(potentialOriginalPath))
                {
                    _logger?.LogInformation($"检测到临时文件，将加载原始文件: {potentialOriginalPath}");
                    filePath = potentialOriginalPath;
                }
            }

            IsLoading = true;
            MethodConfiguration.Clear(); // 清空之前的数据
            _logger?.LogInformation($"正在加载配置文件: {filePath}");
            
            try
            {
                // 使用配置服务加载数据
                var config = await _configService.LoadConfigurationAsync(filePath);
                
                // 确保在UI线程上更新集合
                Application.Current.Dispatcher.Invoke(() =>
                {
                    MethodConfiguration = new ObservableCollection<AnalysisMethodRule>(config);
                });
                
                _logger?.LogInformation("配置文件加载成功，共 {Count} 条规则。", MethodConfiguration.Count);
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, $"加载配置文件时出错: {filePath}");
                MessageBox.Show($"无法加载配置文件: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                MethodConfiguration.Clear();
            }
            finally
            {
                IsLoading = false;
                (SaveConfigurationCommand as RelayCommand)?.RaiseCanExecuteChanged();
            }
        }

        private async Task SaveConfigurationAsync()
        {
            if (SelectedExcelFile == null || string.IsNullOrEmpty(SelectedExcelFile.FullPath))
            {
                MessageBox.Show("请先选择一个分析方法文件", "无法保存", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // 获取原始文件路径
            string originalFilePath = SelectedExcelFile.FullPath;
            string originalFileName = SelectedExcelFile.DisplayName;

            // 记录当前界面上显示的数据，以便验证保存后是否成功
            var originalData = new List<AnalysisMethodRule>();
            foreach (var rule in MethodConfiguration)
            {
                originalData.Add(new AnalysisMethodRule
                {
                    Index = rule.Index,
                    WellPosition = rule.WellPosition,
                    Channel = rule.Channel,
                    SpeciesName = rule.SpeciesName,
                    JudgeFormula = rule.JudgeFormula,
                    ConcentrationFormula = rule.ConcentrationFormula
                });
            }

            IsLoading = true;
            _logger?.LogInformation($"正在保存配置到文件: {originalFilePath}");
            try
            {
                // 确保目录存在
                string? directory = Path.GetDirectoryName(originalFilePath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                // === 日志记录：保存前的数据 ===
                _logger?.LogInformation("--- 保存前 MethodConfiguration 数据快照 ---");
                for (int i = 0; i < Math.Min(5, originalData.Count); i++)
                {
                    var rule = originalData[i];
                    _logger?.LogInformation($"  Index={rule.Index}, Well={rule.WellPosition}, Channel={rule.Channel}, Species={rule.SpeciesName}, Judge='{rule.JudgeFormula}', Conc='{rule.ConcentrationFormula}'");
                }
                if (originalData.Count > 10) _logger?.LogInformation("  ...");
                for (int i = Math.Max(0, originalData.Count - 5); i < originalData.Count; i++)
                {
                     var rule = originalData[i];
                    _logger?.LogInformation($"  Index={rule.Index}, Well={rule.WellPosition}, Channel={rule.Channel}, Species={rule.SpeciesName}, Judge='{rule.JudgeFormula}', Conc='{rule.ConcentrationFormula}'");
                }
                 _logger?.LogInformation("-----------------------------------------");
                // =============================

                // 使用服务保存配置到Excel
                await _configService.SaveConfigurationAsync(originalFilePath, MethodConfiguration);
                _logger?.LogInformation($"成功调用保存服务: {originalFilePath}");

                // --- 简化后的逻辑 --- 
                // 1. 检查文件是否存在 (可选，服务层已有检查)
                // if (!File.Exists(originalFilePath))
                // {
                //     throw new FileNotFoundException("保存后文件不存在，保存可能失败", originalFilePath);
                // }

                // 2. 清理临时文件和备份文件
                try
                {
                    string dirPath = Path.GetDirectoryName(originalFilePath) ?? "";
                    foreach (string tempFile in Directory.GetFiles(dirPath, "*.bak"))
                    {
                        try { File.Delete(tempFile); }
                        catch (Exception exClean) { _logger?.LogWarning(exClean, "删除备份文件时出错: {file}", tempFile); }
                    }
                    foreach (string tempFile in Directory.GetFiles(dirPath, "temp_*"))
                    {
                         try { File.Delete(tempFile); }
                         catch (Exception exClean) { _logger?.LogWarning(exClean, "删除临时文件时出错: {file}", tempFile); }
                    }
                } 
                catch (Exception ex)
                {
                    _logger?.LogWarning(ex, "清理临时文件/备份文件时发生错误");
                }

                // 3. 显示成功消息
                MessageBox.Show("分析方法配置保存成功", "保存成功", MessageBoxButton.OK, MessageBoxImage.Information);

                // 4. (可选) 标记数据为"未修改"状态，如果需要的话
                // MarkAsUnmodified(); 
                
                // --- 原来的复杂刷新逻辑已移除 ---
                // await Task.Delay(500); // 移除
                // var savedData = await _configService.LoadConfigurationAsync(originalFilePath); // 移除
                // _logger?.LogInformation($"原始数据项数: {originalData.Count}, 保存后数据项数: {savedData.Count}"); // 移除
                // if (savedData.Count != originalData.Count) // 移除
                // { // 移除
                //     _logger?.LogWarning($"保存前后数据项数不一致: 原始={originalData.Count}, 保存后={savedData.Count}"); // 移除
                // } // 移除
                // RefreshExcelFiles(); // 移除
                // await Task.Delay(500); // 移除
                // var updatedFile = ExcelFiles.FirstOrDefault(f => f.FullPath.Equals(originalFilePath, StringComparison.OrdinalIgnoreCase)); // 移除
                // if (updatedFile != null) // 移除
                // { // 移除
                //     SelectedExcelFile = null; // 移除
                //     await Task.Delay(100); // 移除
                //     SelectedExcelFile = updatedFile; // 移除
                //     await LoadConfigurationAsync(updatedFile.FullPath); // 移除
                //     _logger?.LogInformation($"已重新加载文件内容: {updatedFile.FullPath}"); // 移除
                // } // 移除
                // else // 移除
                // { // 移除
                //     _logger?.LogWarning($"保存后找不到原文件: {originalFilePath}"); // 移除
                //     SelectedExcelFile = ExcelFiles.FirstOrDefault(); // 移除
                // } // 移除
                // MessageBox.Show("分析方法配置保存成功", "保存成功", MessageBoxButton.OK, MessageBoxImage.Information); // 移动到前面
                // ----------------------------
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, $"保存配置时出错: {originalFilePath}");
                MessageBox.Show($"保存配置失败: {ex.Message}", "保存错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsLoading = false;
            }
        }

        private void ImportFile()
        {
            if (string.IsNullOrEmpty(MethodFolderPath) || !Directory.Exists(MethodFolderPath))
            {
                MessageBox.Show("请先选择一个有效的方法文件夹", "无法导入", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var openFileDialog = new Microsoft.Win32.OpenFileDialog
            {
                Title = "选择要导入的Excel分析方法文件",
                Filter = "Excel文件 (*.xlsx;*.xls)|*.xlsx;*.xls|所有文件 (*.*)|*.*",
                Multiselect = false
            };

            if (openFileDialog.ShowDialog() == true)
            {
                string sourceFilePath = openFileDialog.FileName;
                string fileName = Path.GetFileName(sourceFilePath);
                string destinationFilePath = Path.Combine(MethodFolderPath, fileName);

                // 检查目标位置是否已存在同名文件
                if (File.Exists(destinationFilePath) && 
                    sourceFilePath.ToLower() != destinationFilePath.ToLower()) // 不是同一个文件
                {
                    var result = MessageBox.Show($"文件 {fileName} 在目标文件夹中已存在，是否覆盖？", 
                        "文件已存在", MessageBoxButton.YesNo, MessageBoxImage.Question);
                    
                    if (result != MessageBoxResult.Yes)
                    {
                        return;
                    }
                }

                try
                {
                    // 如果不是同一个文件，则复制
                    if (sourceFilePath.ToLower() != destinationFilePath.ToLower())
                    {
                        File.Copy(sourceFilePath, destinationFilePath, true);
                        _logger?.LogInformation($"已导入文件: {sourceFilePath} 到 {destinationFilePath}");
                    }
                    
                    // 刷新文件列表
                    RefreshExcelFiles();
                    
                    // 选择导入的文件
                    SelectedExcelFile = ExcelFiles.FirstOrDefault(f => f.FullPath.ToLower() == destinationFilePath.ToLower());
                    
                    MessageBox.Show($"文件 {fileName} 导入成功", "导入完成", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    _logger?.LogError(ex, $"导入文件时出错: {sourceFilePath}");
                    MessageBox.Show($"导入文件失败: {ex.Message}", "导入错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void ExportFile()
        {
            if (SelectedExcelFile == null || string.IsNullOrEmpty(SelectedExcelFile.FullPath))
            {
                MessageBox.Show("请先选择一个分析方法文件", "无法导出", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var saveFileDialog = new Microsoft.Win32.SaveFileDialog
            {
                Title = "导出分析方法文件",
                FileName = SelectedExcelFile.DisplayName,
                DefaultExt = Path.GetExtension(SelectedExcelFile.DisplayName),
                Filter = "Excel文件 (*.xlsx)|*.xlsx|Excel 97-2003文件 (*.xls)|*.xls|所有文件 (*.*)|*.*",
                FilterIndex = 1,
                OverwritePrompt = true
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                string destinationFilePath = saveFileDialog.FileName;
                
                try
                {
                    // 复制文件到选择的位置
                    File.Copy(SelectedExcelFile.FullPath, destinationFilePath, true);
                    _logger?.LogInformation($"已导出文件: {SelectedExcelFile.FullPath} 到 {destinationFilePath}");
                    MessageBox.Show($"文件已成功导出到: {destinationFilePath}", "导出完成", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    _logger?.LogError(ex, $"导出文件时出错: {destinationFilePath}");
                    MessageBox.Show($"导出文件失败: {ex.Message}", "导出错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void AddNewRow()
        {
            if (MethodConfiguration == null)
            {
                MethodConfiguration = new ObservableCollection<AnalysisMethodRule>();
            }

            // 确定新行的序号
            int nextIndex = 1;
            if (MethodConfiguration.Count > 0)
            {
                // 找到当前最大序号并加1
                nextIndex = MethodConfiguration.Max(r => r.Index) + 1;
            }

            // 创建新行
            var newRow = new AnalysisMethodRule
            {
                Index = nextIndex,
                WellPosition = "A1", // 默认值
                Channel = "FAM",     // 默认值
                SpeciesName = "请输入",
                JudgeFormula = "[CT]<36",
                ConcentrationFormula = ""
            };

            MethodConfiguration.Add(newRow);
            _logger?.LogInformation($"已添加新行，序号: {nextIndex}");
        }

        private void DeleteSelectedRows()
        {
            // 直接从MethodConfiguration中处理选择删除
            if (MethodConfiguration == null || MethodConfiguration.Count == 0)
            {
                MessageBox.Show("表格中没有数据可以删除", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            // 获取DataGrid实例
            var dataGrid = GetDataGrid();
            if (dataGrid == null)
            {
                MessageBox.Show("无法访问数据表格控件，请联系开发人员", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            // 如果没有选中行，尝试获取当前单元格所在的行
            if (dataGrid.SelectedItems.Count == 0)
            {
                var currentCell = dataGrid.CurrentCell;
                if (currentCell.Item != null && currentCell.Item is AnalysisMethodRule rule)
                {
                    // 提示用户确认删除
                    var result = MessageBox.Show($"确定要删除序号为 {rule.Index} 的行吗？", 
                                                "确认删除", 
                                                MessageBoxButton.YesNo, 
                                                MessageBoxImage.Question);
                    if (result == MessageBoxResult.Yes)
                    {
                        MethodConfiguration.Remove(rule);
                        
                        // 重新编号
                        int newIndex = 1;
                        foreach (var r in MethodConfiguration.OrderBy(r => r.Index))
                        {
                            r.Index = newIndex++;
                        }
                        
                        _logger?.LogInformation($"已删除序号为 {rule.Index} 的行");
                    }
                    return;
                }
                else
                {
                    MessageBox.Show("请先选择要删除的行", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }
            }

            // 有选中的行，询问用户是否确认删除
            var confirmResult = MessageBox.Show($"确定要删除选中的 {dataGrid.SelectedItems.Count} 行数据吗？", 
                                              "确认删除", 
                                              MessageBoxButton.YesNo, 
                                              MessageBoxImage.Question);
            if (confirmResult != MessageBoxResult.Yes)
            {
                return;
            }

            // 收集要删除的项
            var itemsToRemove = new List<AnalysisMethodRule>();
            foreach (var item in dataGrid.SelectedItems)
            {
                if (item is AnalysisMethodRule rule)
                {
                    itemsToRemove.Add(rule);
                }
            }

            // 执行删除操作
            foreach (var rule in itemsToRemove)
            {
                MethodConfiguration.Remove(rule);
            }

            // 重新编号
            int newIndex2 = 1;
            foreach (var rule in MethodConfiguration.OrderBy(r => r.Index))
            {
                rule.Index = newIndex2++;
            }

            _logger?.LogInformation($"已删除 {itemsToRemove.Count} 行数据");
        }

        private void DeleteFile()
        {
            if (SelectedExcelFile == null || string.IsNullOrEmpty(SelectedExcelFile.FullPath))
            {
                MessageBox.Show("请先选择一个要删除的方法文件", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var result = MessageBox.Show($"确定要删除文件 '{SelectedExcelFile.DisplayName}' 吗？\n此操作不可恢复。", 
                                        "确认删除", 
                                        MessageBoxButton.YesNo, 
                                        MessageBoxImage.Warning);
            if (result != MessageBoxResult.Yes)
            {
                return;
            }

            try
            {
                string filePath = SelectedExcelFile.FullPath;
                string fileName = SelectedExcelFile.DisplayName;
                
                // 确保选中的是另一个文件，再删除当前文件
                SelectedExcelFile = ExcelFiles.FirstOrDefault(f => f.FullPath != filePath);
                
                // 删除文件
                File.Delete(filePath);
                
                // 刷新文件列表
                RefreshExcelFiles();
                
                _logger?.LogInformation($"已删除文件: {filePath}");
                MessageBox.Show($"文件 '{fileName}' 已成功删除", "删除成功", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, $"删除文件时出错: {SelectedExcelFile.FullPath}");
                MessageBox.Show($"删除文件失败: {ex.Message}", "删除错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // 获取当前应用程序中的DataGrid实例
        private DataGrid? GetDataGrid()
        {
            if (_dataGrid != null)
                return _dataGrid;

            // 尝试在主窗口的视觉树中查找DataGrid
            var window = Application.Current.MainWindow;
            if (window == null)
                return null;

            _dataGrid = FindDataGrid(window);
            
            // 如果找到了DataGrid，设置事件监听
            if (_dataGrid != null)
            {
                _dataGrid.MouseRightButtonUp += (s, e) => {
                    // 检查点击位置是否在某行上
                    var hitTestResult = VisualTreeHelper.HitTest(_dataGrid, e.GetPosition(_dataGrid));
                    if (hitTestResult != null)
                    {
                        // 向上遍历视觉树找到DataGridRow
                        DependencyObject obj = hitTestResult.VisualHit;
                        while (obj != null && !(obj is DataGridRow))
                        {
                            obj = VisualTreeHelper.GetParent(obj);
                        }
                        
                        // 如果找到了DataGridRow
                        if (obj is DataGridRow row)
                        {
                            // 选中该行
                            _dataGrid.SelectedItem = row.Item;
                        }
                    }
                };
            }
            
            return _dataGrid;
        }

        private DataGrid? FindDataGrid(DependencyObject parent)
        {
            int childCount = VisualTreeHelper.GetChildrenCount(parent);
            for (int i = 0; i < childCount; i++)
            {
                var child = VisualTreeHelper.GetChild(parent, i);
                
                if (child is DataGrid dataGrid)
                {
                    return dataGrid;
                }
                
                var result = FindDataGrid(child);
                if (result != null)
                {
                    return result;
                }
            }
            return null;
        }

        // 提供一个方法，让View可以设置DataGrid引用
        public void SetDataGrid(DataGrid dataGrid)
        {
            _dataGrid = dataGrid;
            _logger?.LogInformation("DataGrid引用已设置");
        }
    }
} 