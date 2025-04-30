using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace MolePCRConvert4WPF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void MenuTemplateFixTool_Click(object sender, RoutedEventArgs e)
        {
            // 打开模板修复工具窗口
            var templateFixWindow = new TemplateFixWindow();
            templateFixWindow.Owner = this;
            templateFixWindow.ShowDialog();
        }

        private void MenuReportDesigner_Click(object sender, RoutedEventArgs e)
        {
            // 打开自定义报告设计器窗口
            var reportDesignerWindow = new Views.ReportDesigner.ReportDesignerView();
            reportDesignerWindow.Owner = this;
            reportDesignerWindow.ShowDialog();
        }

        private void MenuExit_Click(object sender, RoutedEventArgs e)
        {
            // 退出应用程序
            Close();
        }
    }
}