using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using MolePCRConvert4WPF.Core.Models;
using System.IO;
using Microsoft.Win32;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using MolePCRConvert4WPF.App.Services;
using System.Threading.Tasks;
using MolePCRConvert4WPF.Views.ReportDesigner;
using System.Windows.Threading;
using MolePCRConvert4WPF.App.Utils;

namespace MolePCRConvert4WPF.App.ViewModels
{
    public class ReportDesignerViewModel : INotifyPropertyChanged
    {
        #region 属性
        
        private string _templateName;
        public string TemplateName
        {
            get => _templateName;
            set
            {
                if (_templateName != value)
                {
                    _templateName = value;
                    OnPropertyChanged();
                }
            }
        }

        private string _currentFilePath;
        public string CurrentFilePath
        {
            get => _currentFilePath;
            set
            {
                if (_currentFilePath != value)
                {
                    _currentFilePath = value;
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(WindowTitle));
                    OnPropertyChanged(nameof(IsNewButtonEnabled));
                }
            }
        }

        public string WindowTitle => string.IsNullOrEmpty(CurrentFilePath)
            ? "自定义报告设计器 - 新模板"
            : $"自定义报告设计器 - {Path.GetFileName(CurrentFilePath)}";

        private ObservableCollection<TemplateVariable> _availableVariables;
        public ObservableCollection<TemplateVariable> AvailableVariables
        {
            get => _availableVariables;
            set
            {
                if (_availableVariables != value)
                {
                    _availableVariables = value;
                    OnPropertyChanged();
                }
            }
        }

        private TemplateDesignService _templateService;
        private ReportCustomTemplate _currentTemplate;
        private byte[] _templateData;
        
        // 新增：标记模板是否已修改
        private bool _isModified;
        public bool IsModified
        {
            get => _isModified;
            set
            {
                if (_isModified != value)
                {
                    _isModified = value;
                    OnPropertyChanged();
                    
                    // 如果模板被修改，启动自动保存计时器
                    if (value && _autoSaveTimer != null)
                    {
                        _autoSaveTimer.Start();
                    }
                }
            }
        }
        
        // 新增：自动保存计时器
        private DispatcherTimer _autoSaveTimer;
        
        // 新增：控制新建按钮状态
        public bool IsNewButtonEnabled => !string.IsNullOrEmpty(CurrentFilePath) && !IsModified;
        
        // 引用视图
        public ReportDesignerView View { get; set; }

        // 添加默认模板保存路径
        private string DefaultTemplateDirectory => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates");

        #endregion

        #region 命令

        public ICommand NewTemplateCommand { get; }
        public ICommand OpenTemplateCommand { get; }
        public ICommand SaveTemplateCommand { get; }
        public ICommand SaveAsTemplateCommand { get; }
        public ICommand InsertVariableCommand { get; }
        public ICommand InsertDataTableCommand { get; }
        public ICommand PreviewReportCommand { get; }

        #endregion

        public ReportDesignerViewModel()
        {
            try
            {
                // 初始化服务
                _templateService = new TemplateDesignService();
                
                // 初始化变量列表
                InitializeAvailableVariables();
                
                // 初始化命令
                NewTemplateCommand = new RelayCommand(ExecuteNewTemplate, CanExecuteNewTemplate);
                OpenTemplateCommand = new RelayCommand(ExecuteOpenTemplate);
                SaveTemplateCommand = new RelayCommand(ExecuteSaveTemplate, CanExecuteSaveTemplate);
                SaveAsTemplateCommand = new RelayCommand(ExecuteSaveAsTemplate);
                InsertVariableCommand = new RelayCommand<TemplateVariable>(ExecuteInsertVariable, CanExecuteInsertVariable);
                InsertDataTableCommand = new RelayCommand(ExecuteInsertDataTable);
                PreviewReportCommand = new RelayCommand(ExecutePreviewReport, CanPreviewReport);
                
                // 创建默认空模板
                CreateNewTemplate();
                
                // 初始化自动保存计时器
                InitializeAutoSaveTimer();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"初始化报告设计器视图模型时出错: {ex.Message}", "错误", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            }
        }

        // 初始化自动保存计时器
        private void InitializeAutoSaveTimer()
        {
            _autoSaveTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromMinutes(5) // 5分钟自动保存一次
            };
            
            _autoSaveTimer.Tick += (s, e) =>
            {
                if (IsModified && !string.IsNullOrEmpty(CurrentFilePath))
                {
                    // 执行自动保存
                    ExecuteAutoSave();
                }
            };
        }
        
        // 自动保存方法
        private void ExecuteAutoSave()
        {
            try
            {
                if (View != null && !string.IsNullOrEmpty(CurrentFilePath))
                {
                    // 获取当前模板数据
                    _templateData = View.GetTemplateBytes();
                    
                    // 保存到文件
                    File.WriteAllBytes(CurrentFilePath, _templateData);
                    
                    // 更新修改状态
                    IsModified = false;
                }
            }
            catch (Exception ex)
            {
                // 自动保存失败，仅记录日志，不弹出错误
                Console.WriteLine($"自动保存失败: {ex.Message}");
            }
        }

        // 标记模板已修改
        public void MarkAsModified()
        {
            IsModified = true;
        }

        private void InitializeAvailableVariables()
        {
            _availableVariables = new ObservableCollection<TemplateVariable>
            {
                // 系统变量
                new TemplateVariable { 
                    Category = "系统", 
                    Name = "CurrentDate", 
                    DisplayName = "当前日期", 
                    Description = "当前日期 (yyyy-MM-dd)" 
                },
                new TemplateVariable { 
                    Category = "系统", 
                    Name = "CurrentTime", 
                    DisplayName = "当前时间", 
                    Description = "当前时间 (HH:mm:ss)" 
                },
                
                // 样本信息
                new TemplateVariable { 
                    Category = "样本信息", 
                    Name = "SampleID", 
                    DisplayName = "样本ID", 
                    Description = "样本唯一标识符" 
                },
                new TemplateVariable { 
                    Category = "样本信息", 
                    Name = "SampleName", 
                    DisplayName = "样本名称", 
                    Description = "样本名称" 
                },
                new TemplateVariable { 
                    Category = "样本信息", 
                    Name = "SampleType", 
                    DisplayName = "样本类型", 
                    Description = "样本类型" 
                },
                
                // 板信息
                new TemplateVariable { 
                    Category = "板信息", 
                    Name = "PlateID", 
                    DisplayName = "板ID", 
                    Description = "板的唯一标识符" 
                },
                new TemplateVariable { 
                    Category = "板信息", 
                    Name = "PlateName", 
                    DisplayName = "板名称", 
                    Description = "板名称" 
                },
                new TemplateVariable { 
                    Category = "板信息", 
                    Name = "CreationDate", 
                    DisplayName = "创建日期", 
                    Description = "板的创建日期" 
                },
                
                // 结果统计
                new TemplateVariable { 
                    Category = "结果统计", 
                    Name = "TotalSamples", 
                    DisplayName = "总样本数", 
                    Description = "板上的总样本数" 
                },
                new TemplateVariable { 
                    Category = "结果统计", 
                    Name = "PositiveCount", 
                    DisplayName = "阳性数量", 
                    Description = "阳性结果的数量" 
                },
                new TemplateVariable { 
                    Category = "结果统计", 
                    Name = "NegativeCount", 
                    DisplayName = "阴性数量", 
                    Description = "阴性结果的数量" 
                }
            };
        }

        /// <summary>
        /// 创建新模板
        /// </summary>
        public void CreateNewTemplate()
        {
            try
            {
                // 清理当前设计器状态
                _currentTemplate = new ReportCustomTemplate
                {
                    Name = "新模板",
                    CreatedTime = DateTime.Now,
                    LastModifiedTime = DateTime.Now,
                    Description = "新创建的报告模板",
                    ColumnCount = 10,
                    RowCount = 30,
                    Cells = new List<TemplateCellData>()
                };
                
                // 设置基本属性
                TemplateName = _currentTemplate.Name;
                
                // 重要：清空当前文件路径，这样保存时会使用默认路径
                CurrentFilePath = string.Empty;
                
                // 重置其他状态
                IsModified = false;
                
                // 如果视图已存在，重置ReoGrid控件
                if (View?.ReportGrid != null)
                {
                    View.ReportGrid.Reset();
                    View.ReportGrid.CurrentWorksheet.SetCols(10);
                    View.ReportGrid.CurrentWorksheet.SetRows(30);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"创建新模板时出错: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        #region 命令执行方法

        // 添加是否可以执行新建模板的逻辑
        private bool CanExecuteNewTemplate()
        {
            // 如果当前模板未修改或未保存，则允许创建新模板
            return !IsModified;
        }

        private void ExecuteNewTemplate()
        {
            if (IsModified)
            {
                var result = MessageBox.Show(
                    "当前模板已修改，是否保存更改？",
                    "保存确认",
                    MessageBoxButton.YesNoCancel,
                    MessageBoxImage.Question);

                if (result == MessageBoxResult.Cancel)
                {
                    return; // 取消操作
                }

                if (result == MessageBoxResult.Yes)
                {
                    ExecuteSaveTemplate(); // 保存当前模板
                    
                    // 如果保存失败，IsModified仍然为true，中止新建
                    if (IsModified)
                    {
                        return;
                    }
                }
            }
            
            // 创建新模板
            CreateNewTemplate();
            
            // 通知命令状态可能已更改
            CommandManager.InvalidateRequerySuggested();
        }

        private void ExecuteOpenTemplate()
        {
            // 实现打开模板逻辑
            OpenFileDialog openDialog = new OpenFileDialog
            {
                Filter = "报告模板文件 (*.rgf)|*.rgf|Excel文件 (*.xlsx)|*.xlsx|所有文件 (*.*)|*.*",
                Title = "打开报告模板"
            };

            if (openDialog.ShowDialog() == true)
            {
                if (IsModified)
                {
                    MessageBoxResult result = MessageBox.Show("当前模板有未保存的修改，是否保存？", "提示", 
                        MessageBoxButton.YesNoCancel, MessageBoxImage.Question);
                    
                    if (result == MessageBoxResult.Cancel)
                    {
                        return;
                    }
                    
                    if (result == MessageBoxResult.Yes)
                    {
                        ExecuteSaveTemplate();
                    }
                }
                
                // 加载选择的模板文件
                LoadTemplateAsync(openDialog.FileName);
            }
        }

        public async void LoadTemplateAsync(string filePath)
        {
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                MessageBox.Show($"文件不存在: {filePath}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            
            string extension = Path.GetExtension(filePath).ToLower();
            
            if (extension == ".rgf")
            {
                try
                {
                    // 加载ReoGrid格式的模板
                    _templateData = await Task.Run(() => File.ReadAllBytes(filePath));
                    
                    if (_templateData != null && _templateData.Length > 0)
                    {
                        // 创建临时模板对象
                        _currentTemplate = new ReportCustomTemplate
                        {
                            Name = Path.GetFileNameWithoutExtension(filePath),
                            CreatedTime = File.GetCreationTime(filePath),
                            LastModifiedTime = File.GetLastWriteTime(filePath),
                            Description = "从文件加载的模板",
                            ColumnCount = 10, // 实际列数会从模板中加载
                            RowCount = 30     // 实际行数会从模板中加载
                        };
                        
                        TemplateName = _currentTemplate.Name;
                        CurrentFilePath = filePath;
                        IsModified = false;
                        
                        // 加载到ReoGrid控件（通过视图）
                        if (View != null)
                        {
                            try
                            {
                                using (var ms = new MemoryStream(_templateData))
                                {
                                    View.ReportGrid.Load(ms, unvell.ReoGrid.IO.FileFormat.ReoGridFormat);
                                }
                                
                                // 通知命令可执行状态可能改变
                                NewTemplateCommand.NotifyCanExecuteChangedIfNeeded();
                                SaveTemplateCommand.NotifyCanExecuteChangedIfNeeded();
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show($"加载模板到设计器时出错: {ex.Message}", "加载错误", MessageBoxButton.OK, MessageBoxImage.Error);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"读取模板文件时出错: {ex.Message}", "文件读取错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            else if (extension == ".xlsx")
            {
                // 加载Excel格式的模板，提示需要先另存为rgf格式
                var result = MessageBox.Show(
                    "Excel格式的模板不能直接编辑，是否要将其转换为ReoGrid格式？", 
                    "格式转换", 
                    MessageBoxButton.YesNo, 
                    MessageBoxImage.Question);
                
                if (result == MessageBoxResult.Yes)
                {
                    // 转换Excel到ReoGrid格式
                    // 这里需要实现转换逻辑，临时使用创建新模板代替
                    CreateNewTemplate();
                    TemplateName = Path.GetFileNameWithoutExtension(filePath);
                    IsModified = true; // 标记为已修改，以便保存
                }
            }
            else
            {
                MessageBox.Show($"不支持的文件格式: {extension}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ExecuteSaveTemplate()
        {
            try
            {
                if (string.IsNullOrEmpty(CurrentFilePath))
                {
                    // 如果是新模板，自动保存到默认目录
                    string fileName = GetSafeFileName(TemplateName);
                    string templateDir = DefaultTemplateDirectory;
                    
                    // 确保目录存在
                    if (!Directory.Exists(templateDir))
                    {
                        Directory.CreateDirectory(templateDir);
                    }
                    
                    // 生成文件路径
                    string filePath = Path.Combine(templateDir, $"{fileName}.rgf");
                    
                    // 检查文件是否已存在，如果存在则添加时间戳
                    if (File.Exists(filePath))
                    {
                        // 添加时间戳防止文件名冲突
                        string timeStamp = DateTime.Now.ToString("yyyyMMddHHmmss");
                        filePath = Path.Combine(templateDir, $"{fileName}_{timeStamp}.rgf");
                    }
                    
                    CurrentFilePath = filePath;
                }
                
                if (View != null)
                {
                    // 获取当前模板数据
                    _templateData = View.GetTemplateBytes();
                    
                    // 保存到文件
                    File.WriteAllBytes(CurrentFilePath, _templateData);
                    
                    // 更新修改状态
                    IsModified = false;
                    
                    MessageBox.Show($"模板已成功保存到：\n{CurrentFilePath}", "保存成功", MessageBoxButton.OK, MessageBoxImage.Information);
                    
                    // 通知可能的父窗口刷新模板列表
                    NotifyTemplateListRefresh();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"保存模板时出错: {ex.Message}", "保存错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ExecuteSaveAsTemplate()
        {
            SaveFileDialog saveDialog = new SaveFileDialog
            {
                Filter = "报告模板文件 (*.rgf)|*.rgf|Excel文件 (*.xlsx)|*.xlsx",
                Title = "另存报告模板",
                DefaultExt = ".rgf",
                AddExtension = true,
                InitialDirectory = DefaultTemplateDirectory,
                FileName = GetSafeFileName(TemplateName)
            };
            
            if (saveDialog.ShowDialog() == true)
            {
                CurrentFilePath = saveDialog.FileName;
                ExecuteSaveTemplate();
                
                // 保存成功后，通知按钮状态可能发生变化
                NewTemplateCommand.NotifyCanExecuteChangedIfNeeded();
            }
        }
        
        // 获取安全的文件名（去除非法字符）
        private string GetSafeFileName(string fileName)
        {
            if (string.IsNullOrEmpty(fileName))
            {
                fileName = "新模板";
            }
            
            // 移除文件名中的非法字符
            char[] invalidChars = Path.GetInvalidFileNameChars();
            foreach (char c in invalidChars)
            {
                fileName = fileName.Replace(c, '_');
            }
            
            return fileName;
        }
        
        // 通知父窗口刷新模板列表
        private void NotifyTemplateListRefresh()
        {
            try
            {
                // 查找父窗口
                var window = Window.GetWindow(View);
                if (window?.Owner is Window ownerWindow && 
                    ownerWindow.DataContext is ViewModels.ReportTemplateConfigViewModel viewModel)
                {
                    // 调用刷新命令 - 使用安全的方式调用
                    if (viewModel.RefreshCommand != null)
                    {
                        viewModel.RefreshCommand.Execute(null);
                    }
                }
            }
            catch (Exception ex)
            {
                // 记录异常但不中断流程
                Console.WriteLine($"通知刷新模板列表时出错: {ex.Message}");
            }
        }

        private void ExecuteInsertVariable(TemplateVariable variable)
        {
            if (variable != null && View != null)
            {
                View.InsertVariable(variable);
            }
        }

        private void ExecuteInsertDataTable()
        {
            if (View != null)
            {
                View.InsertDataTableMarker();
            }
        }

        private void ExecutePreviewReport()
        {
            MessageBox.Show("预览功能尚未实现", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private bool CanExecuteSaveTemplate()
        {
            // 当模板被修改时，允许保存
            return IsModified;
        }

        private bool CanExecuteInsertVariable(TemplateVariable variable)
        {
            return variable != null && View != null;
        }

        private bool CanPreviewReport()
        {
            return true;
        }

        #endregion
        
        #region INotifyPropertyChanged 实现

        public event PropertyChangedEventHandler PropertyChanged;
        
        protected void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        #endregion
    }
    
    // 命令实现
    public class RelayCommand : ICommand
    {
        private readonly Action _execute;
        private readonly Func<bool> _canExecute;
        
        public RelayCommand(Action execute, Func<bool> canExecute = null)
        {
            _execute = execute ?? throw new ArgumentNullException(nameof(execute));
            _canExecute = canExecute;
        }
        
        public bool CanExecute(object parameter)
        {
            return _canExecute == null || _canExecute();
        }
        
        public void Execute(object parameter)
        {
            _execute();
        }
        
        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }
        
        public void RaiseCanExecuteChanged()
        {
            CommandManager.InvalidateRequerySuggested();
        }
    }
    
    public class RelayCommand<T> : ICommand
    {
        private readonly Action<T> _execute;
        private readonly Predicate<T> _canExecute;
        
        public RelayCommand(Action<T> execute, Predicate<T> canExecute = null)
        {
            _execute = execute ?? throw new ArgumentNullException(nameof(execute));
            _canExecute = canExecute;
        }
        
        public bool CanExecute(object parameter)
        {
            return _canExecute == null || _canExecute((T)parameter);
        }
        
        public void Execute(object parameter)
        {
            _execute((T)parameter);
        }
        
        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }
    }
} 