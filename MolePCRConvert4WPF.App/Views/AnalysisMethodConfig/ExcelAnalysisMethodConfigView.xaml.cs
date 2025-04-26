using MolePCRConvert4WPF.App.ViewModels;
using System.Windows;

namespace MolePCRConvert4WPF.App.Views.AnalysisMethodConfig
{
    /// <summary>
    /// ExcelAnalysisMethodConfigView.xaml 的交互逻辑
    /// </summary>
    public partial class ExcelAnalysisMethodConfigView : System.Windows.Controls.UserControl
    {
        public ExcelAnalysisMethodConfigView()
        {
            InitializeComponent();
            
            // 在控件加载完成后，将DataGrid引用传递给ViewModel
            this.Loaded += ExcelAnalysisMethodConfigView_Loaded;
        }

        private void ExcelAnalysisMethodConfigView_Loaded(object sender, RoutedEventArgs e)
        {
            if (this.DataContext is ExcelAnalysisMethodConfigViewModel viewModel)
            {
                // 将DataGrid的引用传递给ViewModel
                viewModel.SetDataGrid(ConfigDataGrid);
            }
        }
    }
} 