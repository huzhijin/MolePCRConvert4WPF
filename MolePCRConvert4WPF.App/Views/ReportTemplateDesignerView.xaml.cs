using MolePCRConvert4WPF.App.ViewModels;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using unvell.ReoGrid;
using unvell.ReoGrid.Events;
using unvell.ReoGrid.Graphics;
using MolePCRConvert4WPF.App.Utils;
using MolePCRConvert4WPF.Core.Models;

namespace MolePCRConvert4WPF.App.Views
{
    /// <summary>
    /// ReportTemplateDesignerView.xaml 的交互逻辑
    /// </summary>
    public partial class ReportTemplateDesignerView : UserControl
    {
        private readonly ReportTemplateDesignerViewModel? _viewModel;
        
        public ReportTemplateDesignerView()
        {
            InitializeComponent();
            
            Loaded += ReportTemplateDesignerView_Loaded;
            
            _viewModel = DataContext as ReportTemplateDesignerViewModel;
        }
        
        private void ReportTemplateDesignerView_Loaded(object sender, RoutedEventArgs e)
        {
            if (_viewModel == null) return;
            
            // 监听ReoGrid控件内容变化 - 使用CurrentWorksheet的CellDataChanged事件
            ReoGridControl.CurrentWorksheet.CellDataChanged += (s, args) =>
            {
                _viewModel.NotifyDataModified();
            };
            
            // 监听ViewModel模板数据的变化
            _viewModel.PropertyChanged += (s, args) =>
            {
                if (args.PropertyName == nameof(ReportTemplateDesignerViewModel.TemplateData))
                {
                    if (_viewModel.TemplateData != null)
                    {
                        // 加载模板数据到ReoGrid控件
                        using (var ms = new MemoryStream(_viewModel.TemplateData))
                        {
                            ReoGridControl.Load(ms, unvell.ReoGrid.IO.FileFormat.ReoGridFormat);
                        }
                    }
                }
            };
            
            // 监听变量插入请求事件
            if (_viewModel != null)
            {
                _viewModel.VariableInsertRequested += (s, variable) =>
                {
                    InsertVariable(variable);
                };
            }
            
            // 添加快捷键
            PreviewKeyDown += (s, ke) =>
            {
                // Ctrl+S 保存
                if (ke.Key == Key.S && Keyboard.Modifiers == ModifierKeys.Control)
                {
                    if (_viewModel?.SaveCommand.CanExecute(null) == true)
                    {
                        _viewModel.SaveCommand.Execute(null);
                        ke.Handled = true;
                    }
                }
                
                // Ctrl+Shift+S 另存为
                if (ke.Key == Key.S && Keyboard.Modifiers == (ModifierKeys.Control | ModifierKeys.Shift))
                {
                    if (_viewModel?.SaveAsCommand.CanExecute(null) == true)
                    {
                        _viewModel.SaveAsCommand.Execute(null);
                        ke.Handled = true;
                    }
                }
                
                // Ctrl+N 新建
                if (ke.Key == Key.N && Keyboard.Modifiers == ModifierKeys.Control)
                {
                    if (_viewModel?.NewTemplateCommand.CanExecute(null) == true)
                    {
                        _viewModel.NewTemplateCommand.Execute(null);
                        ke.Handled = true;
                    }
                }
            };
        }
        
        /// <summary>
        /// 插入变量到当前选中的单元格
        /// </summary>
        /// <param name="variable">要插入的变量</param>
        private void InsertVariable(TemplateVariable variable)
        {
            if (variable == null) return;
            
            var worksheet = ReoGridControl.CurrentWorksheet;
            var selection = worksheet.SelectionRange;
            
            if (selection == null || selection.IsEmpty)
            {
                MessageBox.Show("请先选择一个单元格", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            
            // 插入变量
            worksheet[selection.Row, selection.Col] = variable.Name;
            
            // 设置变量的样式（蓝色文本）
            worksheet.SetRangeStyles(selection, new WorksheetRangeStyle
            {
                Flag = PlainStyleFlag.TextColor,
                TextColor = SolidColor.Blue
            });
            
            // 通知ViewModel模板已修改
            _viewModel?.NotifyDataModified();
        }
        
        /// <summary>
        /// 获取当前模板数据
        /// </summary>
        public byte[] GetTemplateData()
        {
            using (var ms = new MemoryStream())
            {
                ReoGridControl.Save(ms, unvell.ReoGrid.IO.FileFormat.ReoGridFormat);
                return ms.ToArray();
            }
        }
    }
} 