using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Extensions.Logging;
using MolePCRConvert4WPF.Core.Interfaces;
using MolePCRConvert4WPF.Core.Models;
using Microsoft.Win32;
using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace MolePCRConvert4WPF.App.ViewModels
{
    /// <summary>
    /// 临时导航服务Mock
    /// </summary>
    internal class MockNavigationService : INavigationService
    {
        public TViewModel? GetViewModel<TViewModel>() where TViewModel : class
        {
            // 返回null，在实际生产代码中需要替换为真正的实现
            return null;
        }

        public void NavigateTo<TViewModel>() where TViewModel : class
        {
            // 空实现，在实际生产代码中需要替换为真正的实现
            MessageBox.Show($"导航到{typeof(TViewModel).Name}的功能尚未实现", "开发中", MessageBoxButton.OK, MessageBoxImage.Information);
        }
    }

    /// <summary>
    /// 报告模板配置视图模型
    /// </summary>
    public partial class ReportTemplateConfigViewModel : ObservableObject
    {
        private readonly ILogger<ReportTemplateConfigViewModel> _logger;
        private readonly IReportService _reportService;
        private readonly IReportTemplateDesignerService _designerService;
        private readonly INavigationService _navigationService;

        /// <summary>
        /// 模板集合
        /// </summary>
        [ObservableProperty]
        private ObservableCollection<ReportTemplate> _templates = new();

        /// <summary>
        /// 选中的模板
        /// </summary>
        [ObservableProperty]
        private ReportTemplate? _selectedTemplate;

        /// <summary>
        /// 是否正在加载
        /// </summary>
        [ObservableProperty]
        [NotifyCanExecuteChangedFor(nameof(RefreshCommand))]
        [NotifyCanExecuteChangedFor(nameof(AddTemplateCommand))]
        [NotifyCanExecuteChangedFor(nameof(DeleteTemplateCommand))]
        private bool _isLoading;

        /// <summary>
        /// 状态消息
        /// </summary>
        [ObservableProperty]
        private string? _statusMessage;

        // 命令
        public IAsyncRelayCommand RefreshCommand { get; }
        public IRelayCommand AddTemplateCommand { get; }
        public IRelayCommand DeleteTemplateCommand { get; }
        public IRelayCommand ViewTemplateCommand { get; }
        public IRelayCommand DesignTemplateCommand { get; }
        public IRelayCommand CreateNewTemplateCommand { get; }
        
        public ReportTemplateConfigViewModel(
            ILogger<ReportTemplateConfigViewModel> logger,
            IReportService reportService,
            IReportTemplateDesignerService designerService,
            INavigationService? navigationService = null)
        {
            _logger = logger;
            _reportService = reportService;
            _designerService = designerService;
            _navigationService = navigationService ?? new MockNavigationService();
            
            // 初始化命令
            RefreshCommand = new AsyncRelayCommand(LoadTemplatesAsync, () => !IsLoading);
            AddTemplateCommand = new CommunityToolkit.Mvvm.Input.RelayCommand(AddTemplate, () => !IsLoading);
            DeleteTemplateCommand = new CommunityToolkit.Mvvm.Input.RelayCommand(DeleteTemplate, CanDeleteTemplate);
            ViewTemplateCommand = new CommunityToolkit.Mvvm.Input.RelayCommand(ViewTemplate, () => SelectedTemplate != null);
            DesignTemplateCommand = new CommunityToolkit.Mvvm.Input.RelayCommand(DesignTemplate, () => SelectedTemplate != null);
            CreateNewTemplateCommand = new CommunityToolkit.Mvvm.Input.RelayCommand(CreateNewTemplate, () => !IsLoading);

            // 初始加载
            _ = LoadTemplatesAsync();
        }
        
        private async Task LoadTemplatesAsync()
        {
            if (IsLoading) return;
            
            IsLoading = true;
            StatusMessage = "正在加载报告模板...";
            Templates.Clear();
            
            try
            {
                // 优先加载ReoGrid模板
                var reoGridTemplates = await _designerService.GetReportTemplatesAsync();
                foreach (var template in reoGridTemplates)
                {
                    Templates.Add(template);
                }
                
                // 加载传统Excel模板(为了向后兼容)
                var excelTemplates = await _reportService.GetReportTemplatesAsync();
                foreach (var template in excelTemplates.Where(t => !Templates.Any(rt => rt.FilePath == t.FilePath)))
                {
                    Templates.Add(template);
                }
                
                StatusMessage = $"已加载 {Templates.Count} 个报告模板";
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "加载报告模板时出错");
                StatusMessage = $"加载报告模板失败: {ex.Message}";
                MessageBox.Show($"加载报告模板时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsLoading = false;
            }
        }
        
        private void AddTemplate()
        {
            try
            {
                var openDialog = new OpenFileDialog
                {
                    Filter = "报告模板文件(*.rgf;*.xlsx)|*.rgf;*.xlsx|ReoGrid模板(*.rgf)|*.rgf|Excel文件(*.xlsx)|*.xlsx",
                    Title = "选择报告模板文件",
                    Multiselect = false
                };
                
                if (openDialog.ShowDialog() == true)
                {
                    string templatePath = openDialog.FileName;
                    string extension = Path.GetExtension(templatePath).ToLower();
                    
                    // 复制模板到应用的Templates目录
                    string templateDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates");
                    string fileName = Path.GetFileName(templatePath);
                    string destPath = Path.Combine(templateDir, fileName);
                    
                    // 如果目标文件已存在，添加时间戳
                    if (File.Exists(destPath))
                    {
                        string nameWithoutExt = Path.GetFileNameWithoutExtension(fileName);
                        string ext = Path.GetExtension(fileName);
                        fileName = $"{nameWithoutExt}_{DateTime.Now:yyyyMMddHHmmss}{ext}";
                        destPath = Path.Combine(templateDir, fileName);
                    }
                    
                    // 确保模板目录存在
                    if (!Directory.Exists(templateDir))
                    {
                        Directory.CreateDirectory(templateDir);
                    }
                    
                    // 复制模板文件
                    File.Copy(templatePath, destPath);
                    
                    StatusMessage = $"已添加模板：{fileName}";
                    
                    // 刷新模板列表
                    _ = LoadTemplatesAsync();
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "添加报告模板时出错");
                StatusMessage = $"添加报告模板失败: {ex.Message}";
                MessageBox.Show($"添加报告模板时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        private bool CanDeleteTemplate()
        {
            return !IsLoading && SelectedTemplate != null;
        }
        
        private async void DeleteTemplate()
        {
            if (SelectedTemplate == null) return;
            
            try
            {
                var result = MessageBox.Show(
                    $"确定要删除模板 \"{SelectedTemplate.Name}\" 吗？此操作不可撤销。", 
                    "确认删除", 
                    MessageBoxButton.YesNo, 
                    MessageBoxImage.Question);
                
                if (result == MessageBoxResult.Yes)
                {
                    bool deleted = false;
                    if (SelectedTemplate.IsReoGridTemplate)
                    {
                        deleted = await _designerService.DeleteTemplateAsync(SelectedTemplate);
                    }
                    else
                    {
                        string filePath = SelectedTemplate.FilePath;
                        if (File.Exists(filePath))
                        {
                            File.Delete(filePath);
                            deleted = true;
                        }
                    }
                    
                    if (deleted)
                    {
                        StatusMessage = $"已删除模板：{SelectedTemplate.Name}";
                        
                        // 刷新模板列表
                        await LoadTemplatesAsync();
                    }
                    else
                    {
                        StatusMessage = $"模板文件不存在：{SelectedTemplate.FilePath}";
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "删除报告模板时出错");
                StatusMessage = $"删除报告模板失败: {ex.Message}";
                MessageBox.Show($"删除报告模板时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        private void ViewTemplate()
        {
            if (SelectedTemplate == null) return;
            
            try
            {
                // 使用默认应用打开模板文件
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = SelectedTemplate.FilePath,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "查看报告模板时出错");
                StatusMessage = $"查看报告模板失败: {ex.Message}";
                MessageBox.Show($"查看报告模板时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        private void DesignTemplate()
        {
            if (SelectedTemplate == null) return;
            
            try
            {
                // 创建设计器视图模型并导航到设计器视图
                var designerViewModel = _navigationService.GetViewModel<ReportTemplateDesignerViewModel>();
                if (designerViewModel != null)
                {
                    // 在导航后加载模板
                    _navigationService.NavigateTo<ReportTemplateDesignerViewModel>();
                    
                    // 异步加载模板
                    _ = Task.Run(async () =>
                    {
                        await designerViewModel.LoadTemplateAsync(SelectedTemplate);
                    });
                }
                else
                {
                    throw new InvalidOperationException("无法创建设计器视图模型");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "打开报告模板设计器时出错");
                StatusMessage = $"打开设计器失败: {ex.Message}";
                MessageBox.Show($"打开报告模板设计器时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        private void CreateNewTemplate()
        {
            try
            {
                // 创建设计器视图模型并导航到设计器视图
                var designerViewModel = _navigationService.GetViewModel<ReportTemplateDesignerViewModel>();
                if (designerViewModel != null)
                {
                    // 导航到设计器视图
                    _navigationService.NavigateTo<ReportTemplateDesignerViewModel>();
                    
                    // 执行新建模板命令
                    designerViewModel.NewTemplateCommand.Execute(null);
                }
                else
                {
                    throw new InvalidOperationException("无法创建设计器视图模型");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "创建新模板时出错");
                StatusMessage = $"创建新模板失败: {ex.Message}";
                MessageBox.Show($"创建新模板时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
} 