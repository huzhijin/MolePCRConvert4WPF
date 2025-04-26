using System;
using System.IO;
using System.Windows;
using System.Windows.Media;
using System.Collections.ObjectModel;
using Microsoft.Win32;
using System.Linq;

namespace MolePCRConvert4WPF
{
    public partial class TemplateFixWindow : Window
    {
        public class TemplateFixResult
        {
            public string FileName { get; set; }
            public string Status { get; set; }
            public string Message { get; set; }
            public Brush StatusColor { get; set; }
        }

        private ObservableCollection<TemplateFixResult> _results = new ObservableCollection<TemplateFixResult>();

        public TemplateFixWindow()
        {
            InitializeComponent();
            lvResults.ItemsSource = _results;
        }

        private void BtnFixAll_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 确定模板目录
                string templateDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates");
                if (!Directory.Exists(templateDir))
                {
                    MessageBox.Show($"模板目录不存在: {templateDir}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // 清空结果列表
                _results.Clear();

                // 修复所有模板
                var results = TemplateFixTool.FixAllTemplates(templateDir);
                foreach (var result in results)
                {
                    _results.Add(new TemplateFixResult
                    {
                        FileName = result.FileName,
                        Status = result.Success ? "成功" : "失败",
                        Message = result.Message,
                        StatusColor = result.Success ? Brushes.Green : Brushes.Red
                    });
                }

                MessageBox.Show($"处理完成，共处理 {results.Count} 个模板文件。", "完成", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"执行过程中出错: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnFixSelected_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 创建打开文件对话框
                OpenFileDialog openDialog = new OpenFileDialog
                {
                    Filter = "Excel文件 (*.xlsx)|*.xlsx",
                    Title = "选择要修复的Excel模板",
                    Multiselect = true,
                    InitialDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates")
                };

                // 显示对话框
                if (openDialog.ShowDialog() == true)
                {
                    // 清空结果列表
                    _results.Clear();

                    // 修复选中的模板
                    foreach (string fileName in openDialog.FileNames)
                    {
                        var (success, message) = TemplateFixTool.FixTemplate(fileName);
                        _results.Add(new TemplateFixResult
                        {
                            FileName = Path.GetFileName(fileName),
                            Status = success ? "成功" : "失败",
                            Message = message,
                            StatusColor = success ? Brushes.Green : Brushes.Red
                        });
                    }

                    MessageBox.Show($"处理完成，共处理 {openDialog.FileNames.Length} 个模板文件。", "完成", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"执行过程中出错: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
} 