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
        
        // 命令
        public IRelayCommand SaveCommand { get; }
        public IRelayCommand SaveAsCommand { get; }
        public IRelayCommand NewTemplateCommand { get; }
        
        public ReportTemplateDesignerViewModel(
            ILogger<ReportTemplateDesignerViewModel> logger,
            IReportTemplateDesignerService designerService)
        {
            _logger = logger;
            _designerService = designerService;
            
            // 初始化命令
            SaveCommand = new CommunityToolkit.Mvvm.Input.RelayCommand(SaveTemplate, CanSave);
            SaveAsCommand = new CommunityToolkit.Mvvm.Input.RelayCommand(SaveTemplateAs, () => !IsLoading);
            NewTemplateCommand = new CommunityToolkit.Mvvm.Input.RelayCommand(CreateNewTemplate, () => !IsLoading);
        }
        
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
        /// 是否可以保存
        /// </summary>
        private bool CanSave()
        {
            return !IsLoading && CurrentTemplate != null && IsModified;
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