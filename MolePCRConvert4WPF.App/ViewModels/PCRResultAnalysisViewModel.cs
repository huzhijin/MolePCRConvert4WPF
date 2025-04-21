using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Extensions.Logging;
using MolePCRConvert4WPF.Core.Models;
using MolePCRConvert4WPF.Core.Services;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

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
        private readonly IPCRAnalysisService _pcrAnalysisService;

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
            IPCRAnalysisService pcrAnalysisService)
        {
            _logger = logger;
            _appStateService = appStateService;
            _methodConfigService = methodConfigService;
            _pcrAnalysisService = pcrAnalysisService;

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

                // 使用数据副本调用分析服务，而不是原始数据
                List<AnalysisResultItem> analysisResults = await _pcrAnalysisService.AnalyzeAsync(plateCopy, analysisConfig);

                // Process results for display (sorting, grouping logic)
                ProcessAnalysisResults(analysisResults);

                StatusMessage = $"分析完成. 共 {ResultItems.Count} 条记录, {PatientCount} 个患者.";
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
                    CtValue = result.CtValue,
                    // 处理浓度值 - 使用科学计数法格式化
                    Concentration = result.Concentration.HasValue 
                        ? FormatConcentration(result.Concentration.Value)
                        : "-",
                    // 处理检测结果 - 保留原始结果字符串
                    Result = result.DetectionResult ?? "-",
                    // 判断是否为阳性
                    IsPositive = IsPositiveResult(result.DetectionResult),
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

        // --- Placeholder Command Implementations ---
        private bool CanGenerateReport()
        {
            return ResultItems.Any() && !IsLoading;
        }

        private void GeneratePatientReport()
        {
            StatusMessage = "生成患者报告功能暂未实现.";
            _logger.LogWarning("GeneratePatientReportCommand executed but not implemented.");
            MessageBox.Show("生成患者报告功能暂未实现。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void GeneratePlateReport()
        {
            StatusMessage = "生成整板报告功能暂未实现.";
            _logger.LogWarning("GeneratePlateReportCommand executed but not implemented.");
            MessageBox.Show("生成整板报告功能暂未实现。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void ExportResults()
        {
            StatusMessage = "导出结果功能暂未实现.";
            _logger.LogWarning("ExportCommand executed but not implemented.");
             MessageBox.Show("导出结果功能暂未实现。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
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