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

        // 新增字段：存储最新的原始分析结果
        private List<AnalysisResultItem>? _latestAnalysisResults;

        /// <summary>
        /// 结果项集合
        /// </summary>
        [ObservableProperty]
        private ObservableCollection<PCRResultItem> _resultItems = new();
        
        /// <summary>
        /// 是否有分析结果
        /// </summary>
        [ObservableProperty]
        private bool _hasResults;
        
        /// <summary>
        /// 总记录数
        /// </summary>
        [ObservableProperty]
        private int _totalCount;
        
        /// <summary>
        /// 患者数量
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

            IsLoading = true;
            StatusMessage = "正在加载和分析数据...";
            ResultItems.Clear(); // Clear previous results
            _latestAnalysisResults = null; // Clear previous raw analysis results

            try
            {
                var currentPlate = _appStateService.CurrentPlate;
                // Get the analysis method file path (assuming it's stored in AppStateService)
                string? methodFilePath = _appStateService.CurrentAnalysisMethodPath;

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
                
                // 创建一个当前板的深度副本用于分析，避免修改原始数据
                // 只复制Plate类实际存在的属性
                var plateCopy = new Plate
                {
                    Id = currentPlate.Id,
                    Name = currentPlate.Name,
                    Rows = currentPlate.Rows,
                    Columns = currentPlate.Columns,
                    InstrumentType = currentPlate.InstrumentType,
                    // 移除了不存在的属性：Description, FileName, DateCreated
                    WellLayouts = new List<WellLayout>()
                };
                
                // 复制所有孔位数据
                foreach (var originalWell in currentPlate.WellLayouts)
                {
                    // 对每个孔位创建副本
                    var wellCopy = new WellLayout
                    {
                        Id = originalWell.Id,
                        Row = originalWell.Row,
                        Column = originalWell.Column,
                        // Position是只读属性，会自动计算
                        PatientName = originalWell.PatientName,
                        PatientCaseNumber = originalWell.PatientCaseNumber,
                        SampleName = originalWell.SampleName,
                        CtValue = originalWell.CtValue,
                        Channel = originalWell.Channel
                        // 复制其他需要的属性
                    };
                    plateCopy.WellLayouts.Add(wellCopy);
                }
                
                _logger.LogInformation("Created a deep copy of the plate with {count} wells for analysis", 
                    plateCopy.WellLayouts.Count);

                // Load the configuration rules using the correct service method
                ObservableCollection<AnalysisMethodRule> configRules;
                try
                {
                    configRules = await _methodConfigService.LoadConfigurationAsync(methodFilePath);
                    if (configRules == null || !configRules.Any())
                    {
                        throw new Exception("加载的分析配置规则为空.");
                    }
                }
                catch (Exception configEx)
                {
                    _logger.LogError(configEx, "Failed to load analysis configuration from {FilePath}", methodFilePath);
                    StatusMessage = $"加载分析配置失败: {configEx.Message}";
                    IsLoading = false;
                    return;
                }
                
                // Create an AnalysisMethodConfiguration object from the loaded rules
                // (Assuming PCRAnalysisService expects the wrapper object)
                var analysisConfig = new AnalysisMethodConfiguration
                {
                    Name = System.IO.Path.GetFileNameWithoutExtension(methodFilePath), // Use file name as config name
                    Rules = configRules.ToList()
                };

                // 获取当前仪器类型的PCR分析服务
                IPCRAnalysisService pcrAnalysisService = _pcrAnalysisServiceFactory.GetAnalysisService(currentPlate.InstrumentType);
                _logger.LogInformation($"使用{currentPlate.InstrumentType}类型对应的PCR分析服务");
                
                // 开始分析
                try
                {
                    List<AnalysisResultItem> analysisResults = await pcrAnalysisService.AnalyzeAsync(plateCopy, analysisConfig);
                    if (analysisResults == null || !analysisResults.Any())
                    {
                        _logger.LogWarning("分析结果为空");
                        StatusMessage = "分析结果为空";
                        _latestAnalysisResults = null; // 确保清除
                        IsLoading = false;
                        return;
                    }

                    _logger.LogInformation($"PCR分析完成，获取 {analysisResults.Count} 条结果");
                    
                    // 存储最新的分析结果
                    _latestAnalysisResults = analysisResults;

                    // Process results for display (sorting, grouping logic)
                    ProcessAnalysisResults(analysisResults);

                    StatusMessage = $"分析完成. 共 {ResultItems.Count} 条记录, {PatientCount} 个患者.";
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "分析过程中发生错误");
                    StatusMessage = $"分析出错: {ex.Message}";
                    // Optionally show a message box
                     MessageBox.Show($"分析过程中发生错误:\n{ex.Message}", "分析错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error during data load and analysis.");
                StatusMessage = $"分析出错: {ex.Message}";
                // Optionally show a message box
                 MessageBox.Show($"分析过程中发生错误:\n{ex.Message}", "分析错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsLoading = false;
            }
        }

        private void ProcessAnalysisResults(List<AnalysisResultItem> results)
        {
            if (results == null) return;

            // Sort results: typically by patient first, then well position or channel
            var sortedResults = results
                .OrderBy(r => r.PatientName == "未知患者" ? "Z" : r.PatientName ?? string.Empty) // 将"未知患者"排在最后
                .ThenBy(r => GetRowIndex(r.WellPosition?.Substring(0, 1))) // Sort by Row first (A, B, C...)
                .ThenBy(r => int.TryParse(r.WellPosition?.Substring(1), out int col) ? col : int.MaxValue) // Then by Column (1, 2, 3...)
                .ThenBy(r => r.Channel) // Then by Channel
                .ToList();

            ResultItems.Clear();
            string? lastPatientName = null;
            string? lastCaseNumber = null;
            
            foreach (var result in sortedResults)
            {
                // 确定是否是同一患者的第一行
                bool isFirstPatientRow = result.PatientName != lastPatientName || result.PatientCaseNumber != lastCaseNumber;
                
                if (isFirstPatientRow) 
                {
                    lastPatientName = result.PatientName;
                    lastCaseNumber = result.PatientCaseNumber;
                }
                
                // 创建新的PCRResultItem对象
                var pcrResultItem = new PCRResultItem
                {
                    PatientName = result.PatientName ?? "未知患者",
                    PatientCaseNumber = result.PatientCaseNumber ?? "-",
                    WellPosition = result.WellPosition ?? "-",
                    Channel = result.Channel ?? "-",
                    TargetName = result.TargetName ?? "-",
                    // CT值处理：如果有特殊标记，显示特殊标记；否则显示格式化的CT值
                    CtValue = result.CtValue,
                    CtValueDisplay = !string.IsNullOrEmpty(result.CtValueSpecialMark) 
                        ? result.CtValueSpecialMark 
                        : (result.CtValue.HasValue ? Math.Round(result.CtValue.Value, 2).ToString() : "-"),
                    // 处理浓度值 - 使用科学计数法格式化，如有特殊标记则显示"-"
                    Concentration = !string.IsNullOrEmpty(result.CtValueSpecialMark) 
                        ? "-" 
                        : (result.Concentration.HasValue ? FormatConcentration(result.Concentration.Value) : "-"),
                    // 处理检测结果 - 保留原始结果字符串
                    Result = result.DetectionResult ?? "-",
                    // 判断是否为阳性 - 有特殊标记时不标记为阳性
                    IsPositive = string.IsNullOrEmpty(result.CtValueSpecialMark) && IsPositiveResult(result.DetectionResult),
                    // 添加患者分组标记
                    IsFirstPatientRow = isFirstPatientRow
                };
                
                ResultItems.Add(pcrResultItem);
            }
            
            // 更新统计值
            TotalCount = ResultItems.Count;
            PatientCount = ResultItems
                .Select(r => r.PatientName)
                .Where(name => !string.IsNullOrEmpty(name))
                .Distinct()
                .Count();
            HasResults = TotalCount > 0;
        }

        private bool CanGenerateReport()
        {
            return ResultItems.Any() && !IsLoading;
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
                
                if (!ResultItems.Any())
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
                    foreach (var item in ResultItems)
                    {
                        worksheet.Cells[row, 1].Value = item.PatientName;
                        worksheet.Cells[row, 2].Value = item.PatientCaseNumber;
                        worksheet.Cells[row, 3].Value = item.WellPosition;
                        worksheet.Cells[row, 4].Value = item.Channel;
                        worksheet.Cells[row, 5].Value = item.TargetName;
                        
                        // 设置CT值数据和格式，保留2位小数
                        if (item.CtValue.HasValue)
                        {
                            worksheet.Cells[row, 6].Value = item.CtValue.Value;
                            worksheet.Cells[row, 6].Style.Numberformat.Format = "0.00";
                        }
                        else
                        {
                            worksheet.Cells[row, 6].Value = "-";
                        }
                        
                        worksheet.Cells[row, 7].Value = item.Concentration;
                        worksheet.Cells[row, 8].Value = item.Result;
                        
                        // 如果是阳性结果，设置字体颜色为红色
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
            return !string.IsNullOrEmpty(result) && result.Contains("阳性");
        }
    }
} 