using MolePCRConvert4WPF.App.ViewModels;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

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
            
            // 添加快捷键
            PreviewKeyDown += (s, ke) =>
            {
                // Ctrl+S 保存
                if (ke.Key == Key.S && Keyboard.Modifiers == ModifierKeys.Control)
                {
                    if (_viewModel.SaveCommand.CanExecute(null))
                    {
                        _viewModel.SaveCommand.Execute(null);
                        ke.Handled = true;
                    }
                }
                
                // Ctrl+Shift+S 另存为
                if (ke.Key == Key.S && Keyboard.Modifiers == (ModifierKeys.Control | ModifierKeys.Shift))
                {
                    if (_viewModel.SaveAsCommand.CanExecute(null))
                    {
                        _viewModel.SaveAsCommand.Execute(null);
                        ke.Handled = true;
                    }
                }
                
                // Ctrl+N 新建
                if (ke.Key == Key.N && Keyboard.Modifiers == ModifierKeys.Control)
                {
                    if (_viewModel.NewTemplateCommand.CanExecute(null))
                    {
                        _viewModel.NewTemplateCommand.Execute(null);
                        ke.Handled = true;
                    }
                }
            };
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