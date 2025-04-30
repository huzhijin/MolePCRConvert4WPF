using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Extensions.Logging;
using MolePCRConvert4WPF.Core.Interfaces;
using MolePCRConvert4WPF.Core.Models;
using MolePCRConvert4WPF.Core.Services;
using Microsoft.Win32;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Globalization; // Added for concentration formatting

namespace MolePCRConvert4WPF.App.ViewModels
{
    /// <summary>
    /// PCR结果分析视图模型
    /// </summary>
    public partial class PCRResultAnalysisViewModel : ObservableObject
    {
        private readonly ILogger<PCRResultAnalysisViewModel> _logger;
        private readonly IAppStateService _appStateService;
        private readonly IAnalysisMethodConfigService _methodConfigService; // Assuming this service provides the config
        private readonly IPCRAnalysisServiceFactory _pcrAnalysisServiceFactory;
        private readonly IReportService _reportService;

        // 存储最新的原始分析结果
        private List<AnalysisResultItem>? _latestAnalysisResults;
        // 存储所有定义好的患者和他们的孔位 (来自上一个视图设置)
        private List<WellLayout>? _allDefinedWells;

        /// <summary>
        /// 原始结果项集合（不再直接绑定UI）
        /// </summary>
        //[ObservableProperty] // No longer directly bound
        private ObservableCollection<PCRResultItem> _resultItems = new();

        /// <summary>
        /// 用于绑定到UI的结果集合
        /// </summary>
        [ObservableProperty]
        private ObservableCollection<PCRResultItem> _displayedResults = new();

        /// <summary>
        /// 控制是否显示所有患者（包括无结果的）
        /// </summary>
        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(DisplayedResults))] // Trigger update when this changes
        private bool _showAllPatients = false; // Default to showing only results

        /// <summary>
        /// 是否有分析结果 (基于 DisplayedResults)
        /// </summary>
        [ObservableProperty]
        private bool _hasResults;

        /// <summary>
        /// 总记录数 (基于 DisplayedResults)
        /// </summary>
        [ObservableProperty]
        private int _totalCount;

        /// <summary>
        /// 患者数量 (基于 DisplayedResults)
        /// </summary>
        [ObservableProperty]
        private int _patientCount;

        [ObservableProperty]
        [NotifyPropertyChangedFor(nameof(PatientCount))] // Notify PatientCount changes when ResultEntries changes
        private string? _statusMessage;

        [ObservableProperty]
        [NotifyCanExecuteChangedFor(nameof(RefreshCommand))]
        [NotifyCanExecuteChangedFor(nameof(GeneratePatientReportCommand))]
        [NotifyCanExecuteChangedFor(nameof(GeneratePlateReportCommand))]
        [NotifyCanExecuteChangedFor(nameof(ExportCommand))]
        private bool _isLoading;

        // Commands
        public IAsyncRelayCommand RefreshCommand { get; }
        public IRelayCommand GeneratePatientReportCommand { get; }
        public IRelayCommand GeneratePlateReportCommand { get; }
        public IRelayCommand ExportCommand { get; }

        public PCRResultAnalysisViewModel(
            ILogger<PCRResultAnalysisViewModel> logger,
            IAppStateService appStateService,
            IAnalysisMethodConfigService methodConfigService,
            IPCRAnalysisServiceFactory pcrAnalysisServiceFactory,
            IReportService reportService)
        {
            _logger = logger;
            _appStateService = appStateService;
            _methodConfigService = methodConfigService;
            _pcrAnalysisServiceFactory = pcrAnalysisServiceFactory;
            _reportService = reportService;

            RefreshCommand = new AsyncRelayCommand(LoadAndAnalyzeDataAsync, () => !IsLoading);
            GeneratePatientReportCommand = new CommunityToolkit.Mvvm.Input.RelayCommand(GeneratePatientReport, CanGenerateReport);
            GeneratePlateReportCommand = new CommunityToolkit.Mvvm.Input.RelayCommand(GeneratePlateReport, CanGenerateReport);
            ExportCommand = new CommunityToolkit.Mvvm.Input.RelayCommand(ExportResults, CanGenerateReport);

            // Initial load
            _ = LoadAndAnalyzeDataAsync();
        }

        private async Task LoadAndAnalyzeDataAsync()
        {
            if (IsLoading) return;
            _logger.LogInformation("LoadAndAnalyzeDataAsync started.");

            IsLoading = true;
            StatusMessage = "正在加载和分析数据...";
            // ResultItems.Clear(); // Clear previous results - No longer clearing this directly if needed elsewhere
            DisplayedResults.Clear(); // Clear displayed results
            _latestAnalysisResults = null; // Clear previous raw analysis results
            _allDefinedWells = null; // Clear defined wells

            try
            {
                var currentPlate = _appStateService.CurrentPlate;
                string? methodFilePath = _appStateService.CurrentAnalysisMethodPath;
                _logger.LogInformation("Retrieved CurrentPlate (ID: {PlateId}, Name: {PlateName}) and MethodFilePath: {MethodPath}", currentPlate?.Id, currentPlate?.Name, methodFilePath);

                // <<< ADDED DEBUG LOGGING >>>
                if (currentPlate?.WellLayouts != null)
                {
                    int countWithName = currentPlate.WellLayouts.Count(w => !string.IsNullOrEmpty(w.PatientName));
                    _logger.LogInformation("[DEBUG] Before assigning to _allDefinedWells, currentPlate.WellLayouts has {Count} wells with non-null/empty PatientName.", countWithName);
                    // Log first few patient names found
                    var namesFound = currentPlate.WellLayouts.Where(w => !string.IsNullOrEmpty(w.PatientName)).Select(w => w.PatientName).Distinct().Take(10);
                    _logger.LogInformation("[DEBUG] Names found in currentPlate.WellLayouts: {Names}", string.Join(", ", namesFound));
                }
                else
                {
                     _logger.LogWarning("[DEBUG] currentPlate or currentPlate.WellLayouts is null before assignment check.");
                }
                // <<< END ADDED DEBUG LOGGING >>>

                if (currentPlate == null)
                {
                    StatusMessage = "错误: 未找到当前板数据.";
                    _logger.LogWarning("LoadAndAnalyzeDataAsync: CurrentPlate is null.");
                    IsLoading = false; // Reset loading flag
                    return;
                }
                if (string.IsNullOrEmpty(methodFilePath))
                {
                    StatusMessage = "错误: 未设置分析方法文件路径.";
                    _logger.LogWarning("LoadAndAnalyzeDataAsync: CurrentAnalysisMethodPath is null or empty.");
                    IsLoading = false; // Reset loading flag
                    return;
                }

                // Assign _allDefinedWells field using the potentially updated WellLayouts from AppStateService
                // This should contain the patient names assigned in the previous step.
                if (currentPlate.WellLayouts == null)
                {
                     _logger.LogWarning("currentPlate.WellLayouts is null before assigning _allDefinedWells.");
                     StatusMessage = "错误: 当前板数据缺少孔位布局信息.";
                     IsLoading = false;
                     return;
                }
                _allDefinedWells = currentPlate.WellLayouts.ToList(); // Assign from AppStateService plate's layout
                _logger.LogInformation("Assigned {WellCount} defined wells to _allDefinedWells field (from AppStateService).", _allDefinedWells.Count);

                // Create a deep copy of the plate FOR ANALYSIS, ensuring patient info is copied
                var plateCopy = new Plate
                {
                    Id = currentPlate.Id,
                    Name = currentPlate.Name,
                    Rows = currentPlate.Rows,
                    Columns = currentPlate.Columns,
                    InstrumentType = currentPlate.InstrumentType,
                    WellLayouts = new List<WellLayout>()
                };

                // Copy well data, including patient info from _allDefinedWells (which came from AppStateService)
                foreach (var originalWell in _allDefinedWells) // Iterate over the list we just assigned
                {
                    var wellCopy = new WellLayout
                    {
                        Id = originalWell.Id,
                        Row = originalWell.Row,
                        Column = originalWell.Column,
                        PatientName = originalWell.PatientName, // Copy Name
                        PatientCaseNumber = originalWell.PatientCaseNumber, // Copy CaseNumber
                        SampleName = originalWell.SampleName,
                        CtValue = originalWell.CtValue,
                        Channel = originalWell.Channel
                    };
                    plateCopy.WellLayouts.Add(wellCopy);
                }
                _logger.LogInformation("Created a deep copy of the plate with {count} wells for analysis (including patient info)", plateCopy.WellLayouts.Count);


                // Load the configuration rules using the correct service method
                ObservableCollection<AnalysisMethodRule> configRules;
                try
                {
                    _logger.LogInformation("Loading analysis configuration from: {FilePath}", methodFilePath);
                    configRules = await _methodConfigService.LoadConfigurationAsync(methodFilePath);
                    if (configRules == null || !configRules.Any())
                    {
                        _logger.LogWarning("Loaded analysis configuration rules are null or empty.");
                        throw new Exception("加载的分析配置规则为空.");
                    }
                    _logger.LogInformation("Successfully loaded {RuleCount} analysis configuration rules.", configRules.Count);
                }
                catch (Exception configEx)
                {
                    _logger.LogError(configEx, "Failed to load analysis configuration from {FilePath}", methodFilePath);
                    StatusMessage = $"加载分析配置失败: {configEx.Message}";
                    IsLoading = false;
                    return;
                }

                // Create an AnalysisMethodConfiguration object from the loaded rules
                var analysisConfig = new AnalysisMethodConfiguration
                {
                    Name = System.IO.Path.GetFileNameWithoutExtension(methodFilePath), // Use file name as config name
                    Rules = configRules.ToList()
                };

                // 获取当前仪器类型的PCR分析服务
                IPCRAnalysisService pcrAnalysisService = _pcrAnalysisServiceFactory.GetAnalysisService(currentPlate.InstrumentType);
                _logger.LogInformation($"Using PCR analysis service for instrument type: {currentPlate.InstrumentType}");

                // 开始分析
                try
                {
                    _logger.LogInformation("Calling pcrAnalysisService.AnalyzeAsync with plate copy containing patient info...");
                    List<AnalysisResultItem> analysisResults = await pcrAnalysisService.AnalyzeAsync(plateCopy, analysisConfig);
                    _logger.LogInformation("pcrAnalysisService.AnalyzeAsync completed. Received {ResultCount} results.", analysisResults?.Count ?? 0);

                    // Store the latest raw analysis results
                    _latestAnalysisResults = analysisResults ?? new List<AnalysisResultItem>();

                    // Process results for display
                    ProcessAnalysisResults(); // Initial processing after load

                    // Update status message based on DisplayedResults
                    UpdateStatusMessage(); // Use a helper method to update status
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Error during PCR analysis execution.");
                    StatusMessage = $"分析出错: {ex.Message}";
                    MessageBox.Show($"分析过程中发生错误:\n{ex.Message}", "分析错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Unhandled error during data load and analysis process.");
                StatusMessage = $"分析出错: {ex.Message}";
                MessageBox.Show($"分析过程中发生错误:\n{ex.Message}", "分析错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsLoading = false;
                _logger.LogInformation("LoadAndAnalyzeDataAsync finished.");
            }
        }

        // Renamed ProcessAnalysisResults to handle both modes
        private void ProcessAnalysisResults()
        {
            _logger.LogInformation("ProcessAnalysisResults started. ShowAllPatients: {ShowAll}. LatestResults count: {LatestCount}. DefinedWells count: {DefinedCount}.", ShowAllPatients, _latestAnalysisResults?.Count ?? -1, _allDefinedWells?.Count ?? -1);

            if (_latestAnalysisResults == null || _allDefinedWells == null)
            {
                 _logger.LogWarning("Cannot process results: _latestAnalysisResults or _allDefinedWells is null.");
                 UpdateStatusMessage(); // Update status even if empty
                 DisplayedResults.Clear(); // Ensure display is cleared
                 return; // Nothing to process
            }

            List<PCRResultItem> itemsToAdd = new();
            var patientDataLookup = _allDefinedWells
                                     .Where(w => !string.IsNullOrEmpty(w.PatientName))
                                     .ToLookup(w => w.Position);
            _logger.LogInformation("Created patientDataLookup based on _allDefinedWells. Keys: {KeyCount}", patientDataLookup.Count);

            // 1. Process actual analysis results and attempt to enrich with patient data from _allDefinedWells
            // Also collect the names of patients who definitely have results.
            HashSet<string> patientsWithActualResults = new HashSet<string>();
            if (_latestAnalysisResults != null)
            {
                 _logger.LogInformation("Processing {Count} actual analysis results.", _latestAnalysisResults.Count);
                 foreach (var result in _latestAnalysisResults)
                 {
                    // Primarily trust PatientName from _allDefinedWells based on WellPosition
                    string finalPatientName = "未知患者";
                    string finalCaseNumber = "-";
                    WellLayout? wellInfo = null;

                    if (!string.IsNullOrEmpty(result.WellPosition) && patientDataLookup.Contains(result.WellPosition))
                    {
                        wellInfo = patientDataLookup[result.WellPosition].FirstOrDefault(w => !string.IsNullOrEmpty(w.PatientName));
                    }

                    if (wellInfo != null)
                    {
                        finalPatientName = wellInfo.PatientName; // Trust the definition
                        finalCaseNumber = wellInfo.PatientCaseNumber ?? "-";
                        // Log if the result had a different name (or null)
                        if (result.PatientName != finalPatientName)
                        {
                            _logger.LogInformation("Corrected/Set Patient Name for Well {Well}: '{Name}' (Result was: '{ResultName}')", result.WellPosition, finalPatientName, result.PatientName ?? "NULL");
                        }
                    }
                    else if (!string.IsNullOrEmpty(result.PatientName))
                    {   // Fallback to result's name if lookup failed but result has a name
                        finalPatientName = result.PatientName;
                        finalCaseNumber = result.PatientCaseNumber ?? "-";
                        _logger.LogWarning("Well {Well} not found in patientDataLookup or had no name, using name from result: '{Name}'", result.WellPosition, finalPatientName);
                    }
                    else
                    {   // Truly unknown patient
                        _logger.LogWarning("Could not determine patient name for result at well {Well}. Marking as '未知患者'.", result.WellPosition);
                    }

                    // Add to the set of patients with results IF the name is known
                    if (finalPatientName != "未知患者")
                    {
                        patientsWithActualResults.Add(finalPatientName);
                    }

                    itemsToAdd.Add(new PCRResultItem
                    {
                        PatientName = finalPatientName,
                        PatientCaseNumber = finalCaseNumber, // Use case number from lookup or fallback
                        WellPosition = result.WellPosition,
                        Channel = result.Channel,
                        TargetName = result.TargetName,
                        CtValue = result.CtValue.HasValue ? result.CtValue.Value.ToString("F2") : (result.CtValueSpecialMark ?? "-"),
                        Concentration = result.Concentration.HasValue ? FormatConcentration(result.Concentration.Value) : "-",
                        FinalResult = result.DetectionResult ?? "无判定规则",
                        IsPositive = IsPositiveResult(result.DetectionResult)
                        // IsFirstPatientRow will be set after sorting
                    });
                 }
            }
            _logger.LogInformation("After processing actual results, itemsToAdd count: {Count}. Patients identified with results: {PatientNames}", itemsToAdd.Count, string.Join(", ", patientsWithActualResults));


            // 2. If showing all patients, find defined patients WITHOUT results and add placeholders
            if (ShowAllPatients)
            {
                _logger.LogInformation("ShowAllPatients is true. Processing placeholders...");

                // Get all unique patient info defined in _allDefinedWells (REVISED LOGIC)
                var allDefinedPatientNames = _allDefinedWells
                    .Where(w => !string.IsNullOrEmpty(w.PatientName))
                    .Select(w => w.PatientName)
                    .Distinct()
                    .ToList();
                _logger.LogInformation("Found {Count} unique defined patient names from _allDefinedWells: {PatientNames}", allDefinedPatientNames.Count, string.Join(", ", allDefinedPatientNames));

                // Now reconstruct the patient info list
                var allDefinedPatientsInfo = allDefinedPatientNames
                    .Select(name => {
                        // Find the first well for this patient to get a case number (if any)
                        var wellForPatient = _allDefinedWells.FirstOrDefault(w => w.PatientName == name && !string.IsNullOrEmpty(w.PatientCaseNumber));
                        return new {
                            PatientName = name,
                            CaseNumber = wellForPatient?.PatientCaseNumber ?? "-"
                        };
                    })
                    .ToList();
                 _logger.LogInformation("Reconstructed patient info list with {Count} entries.", allDefinedPatientsInfo.Count);


                int placeholdersAdded = 0;
                foreach (var patientInfo in allDefinedPatientsInfo)
                {
                    // Add placeholder ONLY if this patient name was NOT found among those with actual results
                    if (!patientsWithActualResults.Contains(patientInfo.PatientName))
                    {
                         _logger.LogInformation("--> Adding placeholder for patient defined but without results: {PatientName}, Case#: {CaseNumber}", patientInfo.PatientName, patientInfo.CaseNumber);
                        itemsToAdd.Add(new PCRResultItem
                        {
                            PatientName = patientInfo.PatientName,
                            PatientCaseNumber = patientInfo.CaseNumber,
                            WellPosition = "-", Channel = "-", TargetName = "-", CtValue = "-", Concentration = "-", FinalResult = "无结果", IsPositive = false,
                            IsFirstPatientRow = true // Placeholder row is always the 'first' for that patient display
                        });
                        placeholdersAdded++;
                    } else {
                         _logger.LogInformation("--> Skipping placeholder for patient {PatientName} because they have results.", patientInfo.PatientName);
                    }
                }
                 _logger.LogInformation("Finished processing placeholders. Added {Count} placeholder rows.", placeholdersAdded);
                 _logger.LogInformation("After processing placeholders, itemsToAdd count: {Count}", itemsToAdd.Count);
            }

             // 3. Sort all items (results + placeholders) together
            var sortedResults = itemsToAdd
                .OrderBy(r => r.PatientName == "未知患者" ? 1 : 0) // Put "未知患者" last
                .ThenBy(r => r.PatientName)
                .ThenBy(r => r.PatientCaseNumber) // Sort by case number within patient name
                .ThenBy(r => r.WellPosition == "-" ? 0 : 1) // Put placeholders first within patient group
                .ThenBy(r => GetRowIndex(r.WellPosition?.Length > 0 ? r.WellPosition[0].ToString() : null)) // Sort by Row (A, B, C...)
                .ThenBy(r => int.TryParse(r.WellPosition?.Length > 1 ? r.WellPosition.Substring(1) : string.Empty, out int col) ? col : int.MaxValue) // Then by Column (1, 2, 3...)
                .ThenBy(r => r.Channel == "-" ? 1 : 0) // Keep placeholder channel last (should only apply to placeholder rows)
                .ThenBy(r => r.Channel) // Then by Channel for real results
                .ToList();
             _logger.LogInformation("Sorting completed. {Count} items to display.", sortedResults.Count);

            // 4. Add sorted results to the displayed collection and set IsFirstPatientRow
            DisplayedResults.Clear(); // Clear before adding sorted list
            string? lastPatientName = null;
            foreach (var item in sortedResults)
            {
                bool isFirst = item.PatientName != lastPatientName;

                // Set IsFirstPatientRow based on PatientName change for non-placeholder rows.
                // Placeholder rows already have IsFirstPatientRow=true set during creation.
                if (item.WellPosition != "-") // Only adjust for actual result rows
                {
                    item.IsFirstPatientRow = isFirst;
                }
                // else: keep IsFirstPatientRow = true for placeholders

                DisplayedResults.Add(item);
                lastPatientName = item.PatientName;
            }
             _logger.LogInformation("Finished adding items to DisplayedResults. Final count: {Count}", DisplayedResults.Count);

            UpdateStatusMessage(); // Update count and message
        }

        private void UpdateStatusMessage()
        {
             TotalCount = DisplayedResults.Count;
             PatientCount = DisplayedResults.Select(r => r.PatientName).Distinct().Count();
             HasResults = TotalCount > 0;
             StatusMessage = $"{(ShowAllPatients ? "显示所有患者" : "仅显示有结果的患者")}. 共 {TotalCount} 条记录, {PatientCount} 个患者.";

             // Notify commands about CanExecute changes
             GeneratePatientReportCommand.NotifyCanExecuteChanged();
             GeneratePlateReportCommand.NotifyCanExecuteChanged();
             ExportCommand.NotifyCanExecuteChanged();
        }

        private bool CanGenerateReport()
        {
            // Allow report generation even if showing placeholders, base it on actual analysis results
            // return !IsLoading && _latestAnalysisResults != null && _latestAnalysisResults.Any();
             return !IsLoading && HasResults; // Allow export/report if anything is displayed
        }

        private async void GeneratePatientReport()
        {
            try
            {
                IsLoading = true;
                StatusMessage = "正在准备生成患者报告...";
                
                // 获取当前Plate数据
                var currentPlate = _appStateService.CurrentPlate;
                if (currentPlate == null)
                {
                    MessageBox.Show("无法生成报告：当前没有加载板数据。", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                
                // 获取可用的报告模板
                var templates = await _reportService.GetReportTemplatesAsync();
                
                if (templates == null || !templates.Any())
                {
                    MessageBox.Show("未找到报告模板，请确保模板文件夹中存在有效的模板文件。", "模板缺失", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                
                // 创建选择模板对话框
                var templateSelectionWindow = new Window
                {
                    Title = "选择报告模板",
                    Width = 500,
                    Height = 500, // 增加窗口高度
                    WindowStartupLocation = WindowStartupLocation.CenterScreen,
                    ResizeMode = ResizeMode.NoResize
                };
                
                // 添加ScrollViewer来确保所有内容可见
                var scrollViewer = new ScrollViewer
                {
                    VerticalScrollBarVisibility = ScrollBarVisibility.Auto
                };
                
                var stackPanel = new StackPanel
                {
                    Margin = new Thickness(20)
                };
                
                stackPanel.Children.Add(new TextBlock
                {
                    Text = "请选择生成患者报告使用的模板：",
                    FontSize = 16,
                    Margin = new Thickness(0, 0, 0, 20)
                });
                
                var listBox = new ListBox
                {
                    Height = 250,
                    Margin = new Thickness(0, 0, 0, 20)
                };
                
                foreach (var template in templates)
                {
                    listBox.Items.Add(new ListBoxItem
                    {
                        Content = template.Name,
                        Tag = template
                    });
                }
                
                stackPanel.Children.Add(listBox);
                
                var buttonPanel = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    HorizontalAlignment = HorizontalAlignment.Right
                };
                
                var cancelButton = new Button
                {
                    Content = "取消",
                    Width = 80,
                    Margin = new Thickness(0, 0, 10, 0)
                };
                
                var okButton = new Button
                {
                    Content = "确定",
                    Width = 80,
                    IsEnabled = false
                };
                
                buttonPanel.Children.Add(cancelButton);
                buttonPanel.Children.Add(okButton);
                stackPanel.Children.Add(buttonPanel);
                
                // 设置scrollViewer的内容为stackPanel
                scrollViewer.Content = stackPanel;
                templateSelectionWindow.Content = scrollViewer;
                
                ReportTemplate? selectedTemplate = null;
                
                // 绑定事件
                listBox.SelectionChanged += (sender, e) =>
                {
                    okButton.IsEnabled = listBox.SelectedItem != null;
                };
                
                // 添加双击选择功能
                listBox.MouseDoubleClick += (sender, e) =>
                {
                    if (listBox.SelectedItem != null)
                    {
                        selectedTemplate = (listBox.SelectedItem as ListBoxItem)?.Tag as ReportTemplate;
                        templateSelectionWindow.DialogResult = true;
                    }
                };
                
                cancelButton.Click += (sender, e) =>
                {
                    templateSelectionWindow.DialogResult = false;
                };
                
                okButton.Click += (sender, e) =>
                {
                    if (listBox.SelectedItem is ListBoxItem selectedItem)
                    {
                        selectedTemplate = selectedItem.Tag as ReportTemplate;
                        templateSelectionWindow.DialogResult = true;
                    }
                };
                
                // 显示对话框
                bool? result = templateSelectionWindow.ShowDialog();
                
                if (result != true || selectedTemplate == null)
                {
                    StatusMessage = "用户取消了报告生成";
                    return;
                }
                
                // 先生成预览
                StatusMessage = "正在生成报告预览...";
                // 使用存储的 _latestAnalysisResults
                var (htmlPreviews, patientNames) = await _reportService.GenerateReportPreviewAsync(currentPlate, selectedTemplate, _latestAnalysisResults ?? new List<AnalysisResultItem>(), true);
                
                if (htmlPreviews.Count == 0)
                {
                    MessageBox.Show("未能生成预览，没有找到符合条件的数据。", "预览错误", MessageBoxButton.OK, MessageBoxImage.Warning);
                    StatusMessage = "报告预览生成失败";
                    return;
                }
                
                // 生成临时文件用于导出
                string tempDir = Path.Combine(Path.GetTempPath(), "ReportExports");
                if (!Directory.Exists(tempDir))
                    Directory.CreateDirectory(tempDir);
                
                string tempFilePath = Path.Combine(tempDir, $"PatientReport_{currentPlate.Name}_{DateTime.Now:yyyyMMddHHmmss}.xlsx");
                
                // 在后台生成Excel文件
                // 使用存储的 _latestAnalysisResults
                string reportPath = await _reportService.GenerateExcelReportAsync(currentPlate, selectedTemplate, tempDir, _latestAnalysisResults ?? new List<AnalysisResultItem>(), Path.GetFileName(tempFilePath), true);
                
                // 显示预览窗口
                var previewWindow = new Views.ReportPreviewWindow(htmlPreviews, reportPath, true, patientNames);
                previewWindow.ShowDialog();
                
                StatusMessage = $"患者报告预览已完成";
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "生成患者报告时出错");
                StatusMessage = $"报告生成失败: {ex.Message}";
                MessageBox.Show($"生成患者报告时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsLoading = false;
            }
        }

        private async void GeneratePlateReport()
        {
            try
            {
                IsLoading = true;
                StatusMessage = "正在准备生成整板报告...";
                
                // 获取当前Plate数据
                var currentPlate = _appStateService.CurrentPlate;
                if (currentPlate == null)
                {
                    MessageBox.Show("无法生成报告：当前没有加载板数据。", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                
                // 获取可用的报告模板
                var templates = await _reportService.GetReportTemplatesAsync();
                
                if (templates == null || !templates.Any())
                {
                    MessageBox.Show("未找到报告模板，请确保模板文件夹中存在有效的模板文件。", "模板缺失", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                
                // 创建选择模板对话框
                var templateSelectionWindow = new Window
                {
                    Title = "选择报告模板",
                    Width = 500,
                    Height = 550, // 增加窗口高度
                    WindowStartupLocation = WindowStartupLocation.CenterScreen,
                    ResizeMode = ResizeMode.NoResize
                };
                
                // 添加ScrollViewer来确保所有内容可见
                var scrollViewer = new ScrollViewer
                {
                    VerticalScrollBarVisibility = ScrollBarVisibility.Auto
                };
                
                var stackPanel = new StackPanel
                {
                    Margin = new Thickness(20)
                };
                
                stackPanel.Children.Add(new TextBlock
                {
                    Text = "请选择生成整板报告使用的模板：",
                    FontSize = 16,
                    Margin = new Thickness(0, 0, 0, 20)
                });
                
                var listBox = new ListBox
                {
                    Height = 250,
                    Margin = new Thickness(0, 0, 0, 20)
                };
                
                foreach (var template in templates)
                {
                    listBox.Items.Add(new ListBoxItem
                    {
                        Content = template.Name,
                        Tag = template
                    });
                }
                
                stackPanel.Children.Add(listBox);
                
                var buttonPanel = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    HorizontalAlignment = HorizontalAlignment.Right
                };
                
                var cancelButton = new Button
                {
                    Content = "取消",
                    Width = 80,
                    Margin = new Thickness(0, 0, 10, 0)
                };
                
                var okButton = new Button
                {
                    Content = "确定",
                    Width = 80,
                    IsEnabled = false
                };
                
                buttonPanel.Children.Add(cancelButton);
                buttonPanel.Children.Add(okButton);
                stackPanel.Children.Add(buttonPanel);
                
                // 添加选项让用户选择是生成整板报告还是包含所有患者的报告
                var optionsPanel = new StackPanel
                {
                    Margin = new Thickness(0, 20, 0, 0)
                };
                
                var optionsTitle = new TextBlock
                {
                    Text = "报告类型选项：",
                    FontWeight = FontWeights.Bold,
                    Margin = new Thickness(0, 0, 0, 10)
                };
                
                optionsPanel.Children.Add(optionsTitle);
                
                var standardReportOption = new RadioButton
                {
                    Content = "标准整板报告（简洁视图）",
                    IsChecked = true,
                    Margin = new Thickness(0, 0, 0, 5),
                    Tag = "standard"
                };
                
                var allPatientsOption = new RadioButton
                {
                    Content = "包含所有患者的详细报告（适合多患者数据）",
                    Margin = new Thickness(0, 0, 0, 5),
                    Tag = "all_patients"
                };
                
                optionsPanel.Children.Add(standardReportOption);
                optionsPanel.Children.Add(allPatientsOption);
                stackPanel.Children.Add(optionsPanel);
                
                // 设置scrollViewer的内容为stackPanel
                scrollViewer.Content = stackPanel;
                templateSelectionWindow.Content = scrollViewer;
                
                ReportTemplate? selectedTemplate = null;
                bool isAllPatientsReport = false;
                
                // 绑定事件
                listBox.SelectionChanged += (sender, e) =>
                {
                    okButton.IsEnabled = listBox.SelectedItem != null;
                };
                
                // 添加双击选择功能
                listBox.MouseDoubleClick += (sender, e) =>
                {
                    if (listBox.SelectedItem != null)
                    {
                        selectedTemplate = (listBox.SelectedItem as ListBoxItem)?.Tag as ReportTemplate;
                        isAllPatientsReport = (allPatientsOption.IsChecked == true);
                        templateSelectionWindow.DialogResult = true;
                    }
                };
                
                cancelButton.Click += (sender, e) =>
                {
                    templateSelectionWindow.DialogResult = false;
                };
                
                okButton.Click += (sender, e) =>
                {
                    if (listBox.SelectedItem is ListBoxItem selectedItem)
                    {
                        selectedTemplate = selectedItem.Tag as ReportTemplate;
                        isAllPatientsReport = allPatientsOption.IsChecked == true;
                        templateSelectionWindow.DialogResult = true;
                    }
                };
                
                // 显示对话框
                bool? result = templateSelectionWindow.ShowDialog();
                
                if (result != true || selectedTemplate == null)
                {
                    StatusMessage = "用户取消了报告生成";
                    return;
                }
                
                // 先生成预览
                StatusMessage = "正在生成报告预览...";
                // 使用存储的 _latestAnalysisResults
                var (htmlPreviews, _) = await _reportService.GenerateReportPreviewAsync(currentPlate, selectedTemplate, _latestAnalysisResults ?? new List<AnalysisResultItem>(), false);
                
                if (htmlPreviews.Count == 0)
                {
                    MessageBox.Show("未能生成预览，没有找到符合条件的数据。", "预览错误", MessageBoxButton.OK, MessageBoxImage.Warning);
                    StatusMessage = "报告预览生成失败";
                    return;
                }
                
                // 生成临时文件用于导出
                string tempDir = Path.Combine(Path.GetTempPath(), "ReportExports");
                if (!Directory.Exists(tempDir))
                    Directory.CreateDirectory(tempDir);
                
                string reportType = isAllPatientsReport ? "AllPatients" : "PlateReport";
                string tempFilePath = Path.Combine(tempDir, $"{reportType}_{currentPlate.Name}_{DateTime.Now:yyyyMMddHHmmss}.xlsx");
                
                // 在后台生成Excel文件
                // 使用存储的 _latestAnalysisResults
                string reportPath = await _reportService.GenerateExcelReportAsync(currentPlate, selectedTemplate, tempDir, _latestAnalysisResults ?? new List<AnalysisResultItem>(), Path.GetFileName(tempFilePath), false);
                
                // 显示预览窗口
                var previewWindow = new Views.ReportPreviewWindow(htmlPreviews, reportPath, false);
                previewWindow.ShowDialog();
                
                StatusMessage = $"整板报告预览已完成";
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "生成整板报告时出错");
                StatusMessage = $"报告生成失败: {ex.Message}";
                MessageBox.Show($"生成整板报告时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsLoading = false;
            }
        }

        private void ExportResults()
        {
            try
            {
                IsLoading = true;
                StatusMessage = "正在准备导出结果...";
                
                if (!DisplayedResults.Any())
                {
                    MessageBox.Show("没有可导出的结果数据。", "导出失败", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                
                // 选择保存位置
                var saveDialog = new SaveFileDialog
                {
                    Filter = "Excel文件(*.xlsx)|*.xlsx",
                    Title = "导出PCR结果",
                    FileName = $"PCR分析结果_{DateTime.Now:yyyyMMdd}.xlsx"
                };
                
                if (saveDialog.ShowDialog() != true)
                {
                    StatusMessage = "用户取消了导出";
                    return;
                }
                
                // 创建Excel文件
                StatusMessage = "正在导出结果数据...";
                ExportResultsToExcel(saveDialog.FileName);
                
                StatusMessage = $"结果已导出: {Path.GetFileName(saveDialog.FileName)}";
                
                // 询问是否打开导出的文件
                var result = MessageBox.Show("结果数据已导出，是否立即打开？", "完成", MessageBoxButton.YesNo, MessageBoxImage.Question);
                if (result == MessageBoxResult.Yes)
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = saveDialog.FileName,
                        UseShellExecute = true
                    });
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "导出PCR结果时出错");
                StatusMessage = $"导出失败: {ex.Message}";
                MessageBox.Show($"导出结果时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsLoading = false;
            }
        }
        
        private void ExportResultsToExcel(string filePath)
        {
            // 设置EPPlus许可证上下文
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            
            try
            {
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    // 创建一个工作表
                    var worksheet = package.Workbook.Worksheets.Add("PCR分析结果");
                    
                    // 设置标题行
                    worksheet.Cells[1, 1].Value = "患者姓名";
                    worksheet.Cells[1, 2].Value = "病历号";
                    worksheet.Cells[1, 3].Value = "孔位";
                    worksheet.Cells[1, 4].Value = "通道";
                    worksheet.Cells[1, 5].Value = "靶标名称";
                    worksheet.Cells[1, 6].Value = "CT值";
                    worksheet.Cells[1, 7].Value = "浓度值";
                    worksheet.Cells[1, 8].Value = "检测结果";
                    
                    // 应用标题行样式
                    using (var range = worksheet.Cells[1, 1, 1, 8])
                    {
                        range.Style.Font.Bold = true;
                        range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Medium;
                    }
                    
                    // 填充数据行
                    int row = 2;
                    // Use the DisplayedResults collection which contains ViewModel's PCRResultItem
                    foreach (var item in DisplayedResults)
                    {
                        worksheet.Cells[row, 1].Value = item.PatientName;
                        worksheet.Cells[row, 2].Value = item.PatientCaseNumber; // Use the added property
                        worksheet.Cells[row, 3].Value = item.WellPosition;
                        worksheet.Cells[row, 4].Value = item.Channel;
                        worksheet.Cells[row, 5].Value = item.TargetName;
                        
                        // CT Value is string in ViewModel's PCRResultItem, try parsing back to double for Excel formatting
                        if (double.TryParse(item.CtValue, out double ctValue))
                        {
                            worksheet.Cells[row, 6].Value = ctValue;
                            worksheet.Cells[row, 6].Style.Numberformat.Format = "0.00";
                        }
                        else
                        {
                            worksheet.Cells[row, 6].Value = item.CtValue; // Keep as string if not parseable (e.g., "-")
                        }
                        
                        worksheet.Cells[row, 7].Value = item.Concentration;
                        worksheet.Cells[row, 8].Value = item.FinalResult;
                        
                        // If isPositive is true, set font color to red
                        if (item.IsPositive)
                        {
                            worksheet.Cells[row, 8].Style.Font.Color.SetColor(System.Drawing.Color.Red);
                        }
                        
                        row++;
                    }
                    
                    // 自动调整列宽以适应内容
                    worksheet.Cells.AutoFitColumns();
                    
                    // 保存文件
                    package.Save();
                }
                
                _logger.LogInformation("成功导出PCR结果到Excel: {FilePath}", filePath);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "导出到Excel文件时出错");
                throw;
            }
        }
        
        // Helper to get row index from well position (e.g., 'A' -> 0)
        private int GetRowIndex(string? rowLabel)
        {
            if (string.IsNullOrEmpty(rowLabel) || rowLabel.Length != 1) return -1;
            char label = char.ToUpper(rowLabel[0]);
            return (label >= 'A' && label <= 'Z') ? label - 'A' : -1;
        }

        private string FormatConcentration(double concentration)
        {
            // 使用E4格式，确保显示以科学计数法表示，例如 1.2345E+06
            return concentration.ToString("0.00E+00");
        }

        private bool IsPositiveResult(string? result)
        {
            // Keep this logic based on the result string (which now comes from DetectionResult)
            return !string.IsNullOrEmpty(result) && result.Contains("阳性");
        }

        // Handle property change for ShowAllPatients to update the displayed list
        partial void OnShowAllPatientsChanged(bool value)
        {
            ProcessAnalysisResults(); // Re-process results based on the new flag
        }
    }

    // Ensure PCRResultItem has necessary properties if adding separators etc.
    public partial class PCRResultItem : ObservableObject // Assuming PCRResultItem is defined like this
    {
        [ObservableProperty]
        public string? patientName;
        [ObservableProperty]
        public string? patientCaseNumber;
        [ObservableProperty]
        public string? wellPosition;
        [ObservableProperty]
        public string? channel;
        [ObservableProperty]
        public string? targetName;
        [ObservableProperty]
        public string? ctValue;
        [ObservableProperty]
        public string? concentration;
        [ObservableProperty]
        public string? finalResult;
        [ObservableProperty]
        public bool isPositive; // For row highlighting

        // Added property for visual separation / grouping
        [ObservableProperty]
        public bool isFirstPatientRow;
    }
} 