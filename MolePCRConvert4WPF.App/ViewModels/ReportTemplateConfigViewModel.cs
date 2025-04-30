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
using Microsoft.VisualBasic;
using System.Windows.Input;
using System.Diagnostics;

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
        [NotifyCanExecuteChangedFor(nameof(ViewTemplateCommand))]
        [NotifyCanExecuteChangedFor(nameof(DesignTemplateCommand))]
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
        public IRelayCommand RenameTemplateCommand { get; }
        public IRelayCommand ShowInExplorerCommand { get; }
        
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
            DesignTemplateCommand = new CommunityToolkit.Mvvm.Input.RelayCommand(DesignTemplate, CanDesignTemplate);
            CreateNewTemplateCommand = new CommunityToolkit.Mvvm.Input.RelayCommand(CreateNewTemplate, () => !IsLoading);
            RenameTemplateCommand = new CommunityToolkit.Mvvm.Input.RelayCommand(RenameTemplate, () => SelectedTemplate != null);
            ShowInExplorerCommand = new CommunityToolkit.Mvvm.Input.RelayCommand(ShowInExplorer, () => SelectedTemplate != null);

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
        
        partial void OnSelectedTemplateChanged(ReportTemplate? oldValue, ReportTemplate? newValue)
        {
            try
            {
                // 当选中的模板变化时，通知命令状态更新
                // 使用全局刷新命令状态，避免任何类型转换
                CommandManager.InvalidateRequerySuggested();
                
                // 记录选中的模板
                if (newValue != null)
                {
                    _logger.LogDebug($"已选择模板: {newValue.Name}");
                }
            }
            catch (Exception ex)
            {
                // 捕获并记录任何异常，但不中断操作
                _logger.LogError(ex, "处理模板选择变更时出错");
            }
        }
        
        private bool CanDesignTemplate()
        {
            try
            {
                // 基本检查：如果未选择模板或正在加载，则不能编辑
                if (SelectedTemplate == null || IsLoading)
                {
                    return false;
                }
                
                // 检查文件是否存在
                if (!File.Exists(SelectedTemplate.FilePath))
                {
                    _logger.LogWarning($"模板文件不存在: {SelectedTemplate.FilePath}");
                    return false; // 如果文件不存在，不能设计
                }
                
                // 检查文件扩展名
                string extension = Path.GetExtension(SelectedTemplate.FilePath).ToLowerInvariant();
                
                // 只有ReoGrid模板(.rgf)或Excel模板(.xlsx)才能设计
                bool canDesign = SelectedTemplate.IsReoGridTemplate || 
                                (SelectedTemplate.IsExcelTemplate && extension == ".xlsx");
                
                if (!canDesign)
                {
                    _logger.LogDebug($"模板类型不支持编辑: {SelectedTemplate.Name}, 类型: {extension}");
                }
                
                return canDesign;
            }
            catch (Exception ex)
            {
                // 捕获并记录任何异常，出错时返回false
                _logger.LogError(ex, "检查模板是否可编辑时出错");
                return false;
            }
        }
        
        private void DesignTemplate()
        {
            if (SelectedTemplate == null) return;
            
            ExecuteWithErrorHandling(async () =>
            {
                IsLoading = true;
                StatusMessage = "正在打开模板设计器...";
                
                try
                {
                    if (!File.Exists(SelectedTemplate.FilePath))
                    {
                        MessageBox.Show(
                            $"模板文件不存在: {SelectedTemplate.FilePath}", 
                            "文件缺失", 
                            MessageBoxButton.OK, 
                            MessageBoxImage.Error);
                        return;
                    }
                    
                    if (SelectedTemplate.IsReoGridTemplate)
                    {
                        _logger.LogWarning("编辑 ReoGrid 模板的功能 (OpenTemplateDesignerAsync) 已被禁用，因其接口定义已移除。");
                        await Task.CompletedTask; // 保持方法为 async
                    }
                    else if (SelectedTemplate.IsExcelTemplate)
                    {
                        // 用系统默认程序打开Excel文件
                        _logger.LogInformation($"使用系统默认程序打开Excel模板: {SelectedTemplate.FilePath}");
                        Process.Start(new ProcessStartInfo
                        {
                            FileName = SelectedTemplate.FilePath,
                            UseShellExecute = true
                        });
                    }
                    else
                    {
                        _logger.LogWarning($"不支持编辑的模板类型: {Path.GetExtension(SelectedTemplate.FilePath)}");
                        MessageBox.Show(
                            "不支持编辑此类型的模板文件。", 
                            "不支持的操作", 
                            MessageBoxButton.OK, 
                            MessageBoxImage.Information);
                    }
                    
                    StatusMessage = $"已打开模板: {SelectedTemplate.Name}";
                }
                finally
                {
                    IsLoading = false;
                }
            }, "打开模板设计器");
        }
        
        private void CreateNewTemplate()
        {
            try
            {
                StatusMessage = "正在创建新模板...";
                
                // 打开模板名称输入对话框
                string templateName = "新模板";
                
                string result = Interaction.InputBox(
                    "请输入新模板名称:", 
                    "创建新模板", 
                    templateName);
                
                // 如果用户取消，返回空字符串
                if (string.IsNullOrWhiteSpace(result))
                {
                    return;
                }
                
                templateName = result;
                
                // 创建新模板
                var designerWindow = new MolePCRConvert4WPF.Views.ReportDesigner.ReportDesignerView();
                
                // 设置Owner以支持刷新通知
                if (Application.Current.MainWindow != null)
                {
                    designerWindow.Owner = Application.Current.MainWindow;
                }
                
                var viewModel = designerWindow.DataContext as MolePCRConvert4WPF.App.ViewModels.ReportDesignerViewModel;
                
                if (viewModel != null)
                {
                    viewModel.TemplateName = templateName;
                    // 创建新模板初始状态，确保使用传入的名称
                    viewModel.CreateNewTemplate();
                }
                
                // 显示设计器窗口
                designerWindow.ShowDialog();
                
                // 设计器窗口关闭后刷新模板列表
                _ = LoadTemplatesAsync();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "创建新模板时出错");
                StatusMessage = $"创建新模板失败: {ex.Message}";
                MessageBox.Show($"创建新模板时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        private void RenameTemplate()
        {
            if (SelectedTemplate == null) return;
            
            try
            {
                string currentName = SelectedTemplate.Name;
                _logger.LogInformation($"开始重命名模板: {currentName}");
                
                string result = Interaction.InputBox(
                    "请输入新的模板名称:", 
                    "重命名模板", 
                    currentName);
                
                // 如果用户取消或输入为空，则不进行重命名
                if (string.IsNullOrWhiteSpace(result) || result == currentName)
                {
                    _logger.LogDebug("用户取消重命名或未更改名称");
                    return;
                }
                
                // 检查文件是否存在
                string filePath = SelectedTemplate.FilePath;
                if (!File.Exists(filePath))
                {
                    MessageBox.Show($"模板文件不存在：{filePath}", "文件缺失", MessageBoxButton.OK, MessageBoxImage.Error);
                    _logger.LogWarning($"重命名失败，文件不存在: {filePath}");
                    _ = LoadTemplatesAsync(); // 刷新列表
                    return;
                }
                
                // 更新模板名称
                string oldName = SelectedTemplate.Name;
                SelectedTemplate.Name = result;
                _logger.LogInformation($"模板名称从 '{oldName}' 更改为 '{result}'");
                
                // 更新文件名（如果是ReoGrid模板）
                if (SelectedTemplate.IsReoGridTemplate)
                {
                    try
                    {
                        string directory = Path.GetDirectoryName(filePath);
                        string extension = Path.GetExtension(filePath);
                        
                        // 生成新文件名，确保安全
                        string safeName = GetSafeFileName(result);
                        string newPath = Path.Combine(directory, $"{safeName}{extension}");
                        
                        // 检查新路径是否与原路径相同
                        if (string.Equals(filePath, newPath, StringComparison.OrdinalIgnoreCase))
                        {
                            _logger.LogDebug("新旧文件路径相同，仅更新模板名称");
                            
                            // 仅更新模板对象，不移动文件
                            // 注意：不调用SaveTemplateAsync，避免空引用异常
                            StatusMessage = $"模板已重命名为：{result}";
                            
                            // 直接刷新列表
                            _ = LoadTemplatesAsync();
                            return;
                        }
                        
                        // 检查新文件名是否已存在
                        if (File.Exists(newPath))
                        {
                            // 添加时间戳以避免冲突
                            string timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");
                            newPath = Path.Combine(directory, $"{safeName}_{timestamp}{extension}");
                            _logger.LogDebug($"目标文件已存在，添加时间戳: {newPath}");
                        }
                        
                        _logger.LogInformation($"移动文件: {filePath} -> {newPath}");
                        
                        // 读取原文件内容，以便在SaveTemplateAsync中使用
                        byte[] fileContent = File.ReadAllBytes(filePath);
                        
                        // 移动文件
                        File.Move(filePath, newPath);
                        
                        // 更新模板文件路径
                        SelectedTemplate.FilePath = newPath;
                        
                        StatusMessage = $"模板已重命名为：{result}";
                        
                        // 保存更改（如果服务支持），传递文件内容而非null
                        if (SelectedTemplate.IsReoGridTemplate && fileContent.Length > 0)
                        {
                            _logger.LogDebug($"保存模板变更: {SelectedTemplate.Name}");
                            _ = _designerService.SaveTemplateAsync(SelectedTemplate, fileContent);
                        }
                        
                        // 刷新列表
                        _ = LoadTemplatesAsync();
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError(ex, "重命名模板文件时出错");
                        MessageBox.Show($"重命名模板文件时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
                else
                {
                    // 对于非ReoGrid模板，仅更新名称
                    _logger.LogInformation($"仅更新非ReoGrid模板名称: {SelectedTemplate.Name}");
                    StatusMessage = $"模板名称已更新为：{result}";
                    _ = LoadTemplatesAsync(); // 刷新列表
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "重命名模板时出错");
                StatusMessage = $"重命名模板失败: {ex.Message}";
                MessageBox.Show($"重命名模板时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        private void ShowInExplorer()
        {
            if (SelectedTemplate == null) return;
            
            try
            {
                string filePath = SelectedTemplate.FilePath;
                
                // 检查文件是否存在
                if (!File.Exists(filePath))
                {
                    MessageBox.Show($"模板文件不存在：{filePath}", "文件缺失", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                
                // 打开文件所在目录并选择该文件
                string argument = $"/select,\"{filePath}\"";
                System.Diagnostics.Process.Start("explorer.exe", argument);
                
                StatusMessage = $"已在资源管理器中打开：{filePath}";
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "在资源管理器中显示文件时出错");
                StatusMessage = $"在资源管理器中显示文件失败: {ex.Message}";
                MessageBox.Show($"在资源管理器中显示文件时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        /// <summary>
        /// 生成安全的文件名，移除不允许的字符
        /// </summary>
        private string GetSafeFileName(string fileName)
        {
            if (string.IsNullOrEmpty(fileName)) return "Template";
            
            // 替换所有不允许的文件名字符
            char[] invalidChars = Path.GetInvalidFileNameChars();
            string safeName = fileName;
            
            foreach (char c in invalidChars)
            {
                safeName = safeName.Replace(c, '_');
            }
            
            // 避免文件名过长
            if (safeName.Length > 100)
            {
                safeName = safeName.Substring(0, 100);
            }
            
            // 确保非空
            return string.IsNullOrWhiteSpace(safeName) ? "Template" : safeName;
        }

        private void ExecuteWithErrorHandling(Action action, string operationName)
        {
            if (action == null) return;
            
            try
            {
                _logger.LogDebug($"开始操作: {operationName}");
                action();
                _logger.LogDebug($"完成操作: {operationName}");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"执行{operationName}时出错");
                StatusMessage = $"操作失败: {ex.Message}";
                
                // 向用户显示错误信息
                MessageBox.Show(
                    $"执行{operationName}时出错: {ex.Message}", 
                    "操作错误", 
                    MessageBoxButton.OK, 
                    MessageBoxImage.Error);
            }
            finally
            {
                // 刷新命令状态
                CommandManager.InvalidateRequerySuggested();
            }
        }
    }
} 