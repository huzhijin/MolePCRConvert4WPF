using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Microsoft.Extensions.Logging;
using MolePCRConvert4WPF.Core.Interfaces;
using MolePCRConvert4WPF.Core.Models;
using Microsoft.Win32;
using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Collections.Generic;
using System.Linq;
using unvell.ReoGrid.Graphics;
using unvell.ReoGrid;
using MolePCRConvert4WPF.App.Commands;

namespace MolePCRConvert4WPF.App.ViewModels
{
    /// <summary>
    /// 报告模板设计器视图模型
    /// </summary>
    public partial class ReportTemplateDesignerViewModel : ObservableObject
    {
        private readonly ILogger<ReportTemplateDesignerViewModel> _logger;
        private readonly IReportTemplateDesignerService _designerService;
        
        /// <summary>
        /// 当前编辑的模板
        /// </summary>
        [ObservableProperty]
        private ReportTemplate? _currentTemplate;
        
        /// <summary>
        /// 模板数据
        /// </summary>
        [ObservableProperty]
        private byte[]? _templateData;
        
        /// <summary>
        /// 是否已修改
        /// </summary>
        [ObservableProperty]
        [NotifyCanExecuteChangedFor(nameof(SaveCommand))]
        private bool _isModified;
        
        /// <summary>
        /// 是否正在加载
        /// </summary>
        [ObservableProperty]
        [NotifyCanExecuteChangedFor(nameof(SaveCommand))]
        [NotifyCanExecuteChangedFor(nameof(SaveAsCommand))]
        private bool _isLoading;
        
        /// <summary>
        /// 状态消息
        /// </summary>
        [ObservableProperty]
        private string? _statusMessage;
        
        /// <summary>
        /// 可用的模板变量列表
        /// </summary>
        [ObservableProperty]
        private ObservableCollection<TemplateVariable> _availableVariables = new();
        
        /// <summary>
        /// 按类别分组的模板变量
        /// </summary>
        [ObservableProperty]
        private Dictionary<string, List<MolePCRConvert4WPF.Core.Models.TemplateVariable>> _variablesByCategory = new();
        
        /// <summary>
        /// 按类别分组的变量（用于UI显示）
        /// </summary>
        [ObservableProperty]
        private ObservableCollection<MolePCRConvert4WPF.App.Models.VariableCategoryGroup> _variableCategoryGroups = new();
        
        // 命令
        public CommunityToolkit.Mvvm.Input.IRelayCommand SaveCommand { get; }
        public CommunityToolkit.Mvvm.Input.IRelayCommand SaveAsCommand { get; }
        public CommunityToolkit.Mvvm.Input.IRelayCommand NewTemplateCommand { get; }
        public CommunityToolkit.Mvvm.Input.IRelayCommand<Core.Models.TemplateVariable> InsertVariableCommand { get; }
        
        public ReportTemplateDesignerViewModel(
            ILogger<ReportTemplateDesignerViewModel> logger,
            IReportTemplateDesignerService designerService)
        {
            _logger = logger;
            _designerService = designerService;
            
            // 使用完全限定的命名空间来避免RelayCommand冲突
            SaveCommand = new CommunityToolkit.Mvvm.Input.RelayCommand(SaveTemplate, CanSave);
            SaveAsCommand = new CommunityToolkit.Mvvm.Input.RelayCommand(SaveTemplateAs, () => !IsLoading);
            NewTemplateCommand = new CommunityToolkit.Mvvm.Input.RelayCommand(CreateNewTemplate, () => !IsLoading);
            InsertVariableCommand = new CommunityToolkit.Mvvm.Input.RelayCommand<Core.Models.TemplateVariable>(InsertVariable, (v) => v != null && !IsLoading);
            
            // 初始化变量
            InitializeTemplateVariables();
        }
        
        /// <summary>
        /// 初始化模板变量
        /// </summary>
        private void InitializeTemplateVariables()
        {
            // 系统变量
            var systemVariables = new List<TemplateVariable>
            {
                new TemplateVariable { Name = "${Date}", DisplayName = "当前日期", Category = "系统", Description = "当前日期 (yyyy-MM-dd)" },
                new TemplateVariable { Name = "${Time}", DisplayName = "当前时间", Category = "系统", Description = "当前时间 (HH:mm:ss)" },
                new TemplateVariable { Name = "${DateTime}", DisplayName = "日期时间", Category = "系统", Description = "当前日期和时间 (yyyy-MM-dd HH:mm:ss)" },
                new TemplateVariable { Name = "${UserName}", DisplayName = "用户名", Category = "系统", Description = "当前操作用户名" },
                new TemplateVariable { Name = "${ComputerName}", DisplayName = "计算机名", Category = "系统", Description = "计算机名称" }
            };
            
            // 样本信息变量
            var sampleVariables = new List<TemplateVariable>
            {
                new TemplateVariable { Name = "${SampleCount}", DisplayName = "样本数量", Category = "样本信息", Description = "检测样本总数" },
                new TemplateVariable { Name = "${PositiveSampleCount}", DisplayName = "阳性样本数", Category = "样本信息", Description = "阳性样本数量" },
                new TemplateVariable { Name = "${NegativeSampleCount}", DisplayName = "阴性样本数", Category = "样本信息", Description = "阴性样本数量" },
                new TemplateVariable { Name = "${UndeterminedSampleCount}", DisplayName = "未确定样本数", Category = "样本信息", Description = "无法确定结果的样本数量" }
            };
            
            // 板信息变量
            var plateVariables = new List<TemplateVariable>
            {
                new TemplateVariable { Name = "${PlateId}", DisplayName = "板ID", Category = "板信息", Description = "检测板唯一标识" },
                new TemplateVariable { Name = "${RunDate}", DisplayName = "运行日期", Category = "板信息", Description = "检测运行日期" },
                new TemplateVariable { Name = "${RunTime}", DisplayName = "运行时间", Category = "板信息", Description = "检测运行时间" },
                new TemplateVariable { Name = "${RunBy}", DisplayName = "操作人员", Category = "板信息", Description = "执行检测的操作人员" },
                new TemplateVariable { Name = "${InstrumentName}", DisplayName = "仪器名称", Category = "板信息", Description = "使用的仪器名称" },
                new TemplateVariable { Name = "${InstrumentSerialNumber}", DisplayName = "仪器序列号", Category = "板信息", Description = "仪器序列号" }
            };
            
            // 检测结果统计变量
            var resultVariables = new List<TemplateVariable>
            {
                new TemplateVariable { Name = "${PositiveRate}", DisplayName = "阳性率", Category = "结果统计", Description = "阳性样本比例 (%)" },
                new TemplateVariable { Name = "${NegativeRate}", DisplayName = "阴性率", Category = "结果统计", Description = "阴性样本比例 (%)" },
                new TemplateVariable { Name = "${UndeterminedRate}", DisplayName = "未确定率", Category = "结果统计", Description = "未确定样本比例 (%)" }
            };
            
            // 合并所有变量到一个列表
            var allVariables = new List<TemplateVariable>();
            allVariables.AddRange(systemVariables);
            allVariables.AddRange(sampleVariables);
            allVariables.AddRange(plateVariables);
            allVariables.AddRange(resultVariables);
            
            // 更新可用变量集合
            AvailableVariables = new ObservableCollection<TemplateVariable>(allVariables.OrderBy(v => v.Category).ThenBy(v => v.DisplayName));
            
            // 按类别分组变量
            VariablesByCategory = allVariables
                .GroupBy(v => v.Category)
                .ToDictionary(g => g.Key, g => g.ToList());
                
            // 转换为UI显示用的分组集合
            VariableCategoryGroups = MolePCRConvert4WPF.App.Utils.TemplateVariableConverter.ConvertDictionaryToVariableCategoryGroupsLinq(VariablesByCategory);
        }
        
        /// <summary>
        /// 插入变量到模板
        /// </summary>
        /// <param name="variable">要插入的变量</param>
        private void InsertVariable(Core.Models.TemplateVariable variable)
        {
            if (variable == null) return;
            
            // 这里需要实现ReoGrid控件的变量插入逻辑
            // 通常是将变量名称插入到当前选中的单元格中
            
            // 产生一个事件，让视图知道需要插入变量
            VariableInsertRequested?.Invoke(this, variable);
            
            // 标记为已修改
            IsModified = true;
        }
        
        /// <summary>
        /// 变量插入请求事件
        /// </summary>
        public event EventHandler<Core.Models.TemplateVariable> VariableInsertRequested;
        
        /// <summary>
        /// 加载模板
        /// </summary>
        public async Task LoadTemplateAsync(ReportTemplate template)
        {
            if (IsLoading) return;
            
            IsLoading = true;
            StatusMessage = $"正在加载模板: {template.Name}...";
            
            try
            {
                TemplateData = await _designerService.LoadTemplateAsync(template);
                CurrentTemplate = template;
                IsModified = false;
                
                StatusMessage = $"已加载模板: {template.Name}";
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "加载模板时出错");
                StatusMessage = $"加载模板失败: {ex.Message}";
                MessageBox.Show($"加载模板时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsLoading = false;
            }
        }
        
        /// <summary>
        /// 保存模板
        /// </summary>
        private async void SaveTemplate()
        {
            if (CurrentTemplate == null || TemplateData == null) return;
            
            IsLoading = true;
            StatusMessage = $"正在保存模板: {CurrentTemplate.Name}...";
            
            try
            {
                await _designerService.SaveTemplateAsync(CurrentTemplate, TemplateData);
                IsModified = false;
                
                StatusMessage = $"已保存模板: {CurrentTemplate.Name}";
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "保存模板时出错");
                StatusMessage = $"保存模板失败: {ex.Message}";
                MessageBox.Show($"保存模板时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsLoading = false;
            }
        }
        
        /// <summary>
        /// 另存为新模板
        /// </summary>
        private async void SaveTemplateAs()
        {
            if (TemplateData == null) return;
            
            try
            {
                // 弹出输入对话框，让用户输入新模板名称
                var inputDialog = new InputDialog("模板名称", "请输入新模板名称:", "新模板");
                if (inputDialog.ShowDialog() != true)
                {
                    return;
                }
                
                string newTemplateName = inputDialog.Input;
                if (string.IsNullOrWhiteSpace(newTemplateName))
                {
                    MessageBox.Show("模板名称不能为空！", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                
                IsLoading = true;
                StatusMessage = $"正在另存为新模板: {newTemplateName}...";
                
                // 创建新模板
                var newTemplate = await _designerService.CreateNewTemplateAsync(newTemplateName);
                
                // 保存当前数据到新模板
                await _designerService.SaveTemplateAsync(newTemplate, TemplateData);
                
                // 切换到新模板
                CurrentTemplate = newTemplate;
                IsModified = false;
                
                StatusMessage = $"已另存为新模板: {newTemplate.Name}";
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "另存为新模板时出错");
                StatusMessage = $"另存为新模板失败: {ex.Message}";
                MessageBox.Show($"另存为新模板时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsLoading = false;
            }
        }
        
        /// <summary>
        /// 创建新模板
        /// </summary>
        private async void CreateNewTemplate()
        {
            try
            {
                // 如果当前有未保存的更改，提示用户保存
                if (IsModified)
                {
                    var result = MessageBox.Show(
                        "当前模板有未保存的更改，是否保存？",
                        "保存更改",
                        MessageBoxButton.YesNoCancel,
                        MessageBoxImage.Question);
                        
                    if (result == MessageBoxResult.Cancel)
                    {
                        return;
                    }
                    
                    if (result == MessageBoxResult.Yes)
                    {
                        SaveTemplate();
                    }
                }
                
                // 弹出输入对话框，让用户输入新模板名称
                var inputDialog = new InputDialog("新建模板", "请输入新模板名称:", "新模板");
                if (inputDialog.ShowDialog() != true)
                {
                    return;
                }
                
                string newTemplateName = inputDialog.Input;
                if (string.IsNullOrWhiteSpace(newTemplateName))
                {
                    MessageBox.Show("模板名称不能为空！", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                
                IsLoading = true;
                StatusMessage = $"正在创建新模板: {newTemplateName}...";
                
                // 创建新模板
                var newTemplate = await _designerService.CreateNewTemplateAsync(newTemplateName);
                
                // 加载新模板
                await LoadTemplateAsync(newTemplate);
                
                StatusMessage = $"已创建新模板: {newTemplate.Name}";
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "创建新模板时出错");
                StatusMessage = $"创建新模板失败: {ex.Message}";
                MessageBox.Show($"创建新模板时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsLoading = false;
            }
        }
        
        /// <summary>
        /// 通知模板数据已修改
        /// </summary>
        public void NotifyDataModified()
        {
            IsModified = true;
        }
        
        /// <summary>
        /// 更新模板数据
        /// </summary>
        public void UpdateTemplateData(byte[] data)
        {
            TemplateData = data;
            IsModified = true;
        }
        
        /// <summary>
        /// 检查是否可以保存
        /// </summary>
        private bool CanSave()
        {
            return !IsLoading && CurrentTemplate != null && IsModified;
        }
        
        /// <summary>
        /// 手动通知命令状态已变更
        /// </summary>
        public void NotifyCommandsCanExecuteChanged()
        {
            // 使用CommunityToolkit.Mvvm.Input中的NotifyCanExecuteChanged方法
            (SaveCommand as CommunityToolkit.Mvvm.Input.IRelayCommand)?.NotifyCanExecuteChanged();
            (SaveAsCommand as CommunityToolkit.Mvvm.Input.IRelayCommand)?.NotifyCanExecuteChanged();
            (NewTemplateCommand as CommunityToolkit.Mvvm.Input.IRelayCommand)?.NotifyCanExecuteChanged();
            (InsertVariableCommand as CommunityToolkit.Mvvm.Input.IRelayCommand)?.NotifyCanExecuteChanged();
        }
    }
    
    /// <summary>
    /// 输入对话框
    /// </summary>
    public class InputDialog : Window
    {
        private System.Windows.Controls.TextBox? inputTextBox;
        
        public string Input { get; private set; } = string.Empty;
        
        public InputDialog(string title, string prompt, string defaultValue = "")
        {
            Title = title;
            Width = 400;
            Height = 180;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            ResizeMode = ResizeMode.NoResize;
            
            var grid = new System.Windows.Controls.Grid();
            grid.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = new GridLength(0, System.Windows.GridUnitType.Auto) });
            grid.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = new GridLength(1, System.Windows.GridUnitType.Star) });
            grid.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = new GridLength(0, System.Windows.GridUnitType.Auto) });
            
            var promptTextBlock = new System.Windows.Controls.TextBlock
            {
                Text = prompt,
                Margin = new Thickness(10),
                TextWrapping = TextWrapping.Wrap
            };
            grid.Children.Add(promptTextBlock);
            System.Windows.Controls.Grid.SetRow(promptTextBlock, 0);
            
            inputTextBox = new System.Windows.Controls.TextBox
            {
                Text = defaultValue,
                Margin = new Thickness(10),
                VerticalAlignment = VerticalAlignment.Center
            };
            grid.Children.Add(inputTextBox);
            System.Windows.Controls.Grid.SetRow(inputTextBox, 1);
            
            var buttonPanel = new System.Windows.Controls.StackPanel
            {
                Orientation = System.Windows.Controls.Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(10)
            };
            
            var okButton = new System.Windows.Controls.Button
            {
                Content = "确定",
                IsDefault = true,
                Padding = new Thickness(20, 5, 20, 5),
                Margin = new Thickness(5)
            };
            okButton.Click += (s, e) => 
            {
                Input = inputTextBox?.Text ?? string.Empty;
                DialogResult = true;
            };
            
            var cancelButton = new System.Windows.Controls.Button
            {
                Content = "取消",
                IsCancel = true,
                Padding = new Thickness(20, 5, 20, 5),
                Margin = new Thickness(5)
            };
            
            buttonPanel.Children.Add(okButton);
            buttonPanel.Children.Add(cancelButton);
            
            grid.Children.Add(buttonPanel);
            System.Windows.Controls.Grid.SetRow(buttonPanel, 2);
            
            Content = grid;
        }
        
        protected override void OnContentRendered(EventArgs e)
        {
            base.OnContentRendered(e);
            inputTextBox?.Focus();
            inputTextBox?.SelectAll();
        }
    }
} 