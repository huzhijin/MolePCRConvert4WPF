using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using Microsoft.Win32;
using MolePCRConvert4WPF.Core.Models;
using System.Windows.Controls;
using System.Diagnostics;
using System.Text;

namespace MolePCRConvert4WPF.App.Views
{
    /// <summary>
    /// ReportPreviewWindow.xaml 的交互逻辑
    /// </summary>
    public partial class ReportPreviewWindow : Window
    {
        private readonly List<string> _htmlReports = new List<string>();
        private string _excelFilePath;
        private int _currentPage = 0;
        private bool _isPatientReport = false;
        private readonly List<string> _patientNames = new List<string>();

        /// <summary>
        /// 初始化报告预览窗口
        /// </summary>
        /// <param name="htmlReports">HTML格式的报告内容列表，每一个表示一页报告</param>
        /// <param name="excelFilePath">Excel文件路径，用于导出</param>
        /// <param name="isPatientReport">是否是患者报告</param>
        /// <param name="patientNames">患者名称列表，与htmlReports对应</param>
        public ReportPreviewWindow(List<string> htmlReports, string excelFilePath, bool isPatientReport, List<string> patientNames = null)
        {
            InitializeComponent();
            
            _htmlReports = htmlReports ?? new List<string>();
            _excelFilePath = excelFilePath;
            _isPatientReport = isPatientReport;
            _patientNames = patientNames ?? new List<string>();
            
            // 设置报告类型显示
            RunReportType.Text = isPatientReport ? "患者报告" : "整板报告";
            
            // 如果没有报告内容，显示提示
            if (_htmlReports.Count == 0 || 
                (_htmlReports.Count > 0 && (_htmlReports[0].Contains("startIndex cannot be larger") || 
                                          _htmlReports[0].Contains("生成表格预览时发生错误"))))
            {
                // 清空现有内容，显示友好的引导信息
                _htmlReports.Clear();
                _htmlReports.Add(@"<html><body>
                    <div style='font-family:Arial; margin:20px; text-align:center;'>
                    <h2 style='color:#5a5a5a;margin-top:50px;'>预览功能暂时不可用</h2>
                    <p style='color:#666;font-size:16px;'>当前模板无法正确生成HTML预览，请使用下方的""导出Excel""按钮。</p>
                    <div style='margin:40px;padding:20px;background-color:#f5f5f5;border-radius:10px;'>
                    <p style='font-weight:bold;color:#333;'>推荐操作步骤：</p>
                    <p style='text-align:left;margin-left:120px;'>
                    1. 点击左下角的""<span style='color:blue;font-weight:bold;'>导出Excel</span>""按钮<br>
                    2. 选择保存位置<br>
                    3. 在Excel中查看和打印报告
                    </p>
                    </div>
                    <p style='color:#888;font-size:14px;margin-top:40px;'>* 系统已自动创建Excel报告文件，内容与模板格式相同</p>
                    </div></body></html>");
            }
            
            // 调整窗口大小，非患者报告时设置更大的窗口
            if (!isPatientReport)
            {
                this.Width = 1000;
                this.Height = 800;
            }
            
            // 初始化页面计数
            UpdatePageInfo();
            
            // 显示第一页
            ShowCurrentPage();
            
            // 窗口加载完成后的处理
            Loaded += (s, e) => {
                // 确保WebBrowser控件正确加载
                UpdatePageInfo();
                ShowCurrentPage();
            };
            
            // 添加WebBrowser脚本错误处理以防止JS错误弹窗
            WebPreview.Navigated += WebPreview_Navigated;
        }

        private void WebPreview_Navigated(object sender, System.Windows.Navigation.NavigationEventArgs e)
        {
            try
            {
                // 设置WebBrowser的脚本错误处理
                dynamic doc = WebPreview.Document;
                if (doc != null)
                {
                    // 启用设计模式以便编辑内容
                    doc.designMode = "On";
                    
                    // 确保documentElement存在再操作style
                    if (doc.documentElement != null && doc.documentElement.style != null)
                    {
                        // 添加额外的空值检查，确保style对象存在
                        try {
                            doc.documentElement.style.overflow = "auto";
                        }
                        catch (Exception styleEx)
                        {
                            Debug.WriteLine($"设置样式时出错: {styleEx.Message}");
                            // 出错时不抛出异常，继续执行
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // 只记录异常，不中断用户体验
                Debug.WriteLine($"WebBrowser设置样式时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 更新页面信息显示
        /// </summary>
        private void UpdatePageInfo()
        {
            RunCurrentPage.Text = (_currentPage + 1).ToString();
            RunTotalPages.Text = _htmlReports.Count.ToString();
            
            // 更新患者信息（仅对患者报告）
            if (_isPatientReport && _patientNames.Count > _currentPage)
            {
                RunPatientName.Text = _patientNames[_currentPage];
                TxtPatientInfo.Visibility = Visibility.Visible;
            }
            else
            {
                TxtPatientInfo.Visibility = Visibility.Collapsed;
            }
            
            // 对于整板报告，如果只有一页，隐藏分页控件
            if (!_isPatientReport && _htmlReports.Count == 1)
            {
                BtnPrevPage.Visibility = Visibility.Collapsed;
                BtnNextPage.Visibility = Visibility.Collapsed;
            }
            else
            {
                // 更新翻页按钮状态
                BtnPrevPage.IsEnabled = _currentPage > 0;
                BtnNextPage.IsEnabled = _currentPage < _htmlReports.Count - 1;
                
                BtnPrevPage.Visibility = Visibility.Visible;
                BtnNextPage.Visibility = Visibility.Visible;
            }
        }

        /// <summary>
        /// 显示当前页的报告内容
        /// </summary>
        private void ShowCurrentPage()
        {
            if (_htmlReports.Count > _currentPage)
            {
                try
                {
                    // 检查HTML内容是否为空或包含错误信息
                    string currentHtml = _htmlReports[_currentPage];
                    if (string.IsNullOrEmpty(currentHtml))
                    {
                        _htmlReports[_currentPage] = @"<html><body><h2 style='color:red;text-align:center;margin-top:100px;'>预览内容为空</h2><p style='text-align:center'>请使用导出Excel功能查看完整报告</p></body></html>";
                    }
                    // 如果出现预览错误，自动尝试导出
                    else if (currentHtml.Contains("startIndex cannot be larger") || 
                             currentHtml.Contains("生成表格预览时发生错误"))
                    {
                        _htmlReports[_currentPage] = @"<html><body>
                            <div style='font-family:Arial; margin:20px; text-align:center;'>
                            <h2 style='color:#5a5a5a;margin-top:50px;'>预览功能暂时不可用</h2>
                            <p style='color:#666;font-size:16px;'>当前模板无法正确生成HTML预览，请使用下方的""导出Excel""按钮。</p>
                            <div style='margin:40px;padding:20px;background-color:#f5f5f5;border-radius:10px;'>
                            <p style='font-weight:bold;color:#333;'>提示：</p>
                            <p style='text-align:left;margin-left:120px;'>
                            - 部分图表和图片只能在Excel中正确显示<br>
                            - 复杂格式可能在预览中无法完全呈现<br>
                            - 请使用Excel查看完整报告格式
                            </p>
                            </div>
                            <p style='color:#888;font-size:14px;margin-top:40px;'>* 系统已自动创建Excel报告文件，内容与模板格式相同</p>
                            </div></body></html>";
                    }

                    // 创建临时HTML文件
                    string tempFile = Path.Combine(Path.GetTempPath(), $"report_preview_{Guid.NewGuid()}.html");
                    
                    // 增强HTML，添加额外的图片支持和样式
                    string enhancedHtml = EnhanceHtmlWithImageSupport(currentHtml);
                    
                    File.WriteAllText(tempFile, enhancedHtml);
                    
                    // 使用WebBrowser显示
                    WebPreview.Navigate(new Uri(tempFile));
                    
                    // 更新页面信息
                    UpdatePageInfo();
                }
                catch (Exception ex)
                {
                    // 记录异常详情
                    string errorDetails = $"显示预览时出错: {ex.Message}\n调用栈: {ex.StackTrace}";
                    Debug.WriteLine(errorDetails);
                    
                    // 创建错误HTML并显示
                    try
                    {
                        string errorHtml = @"<html><body><h2 style='color:red;text-align:center;margin-top:50px;'>预览显示错误</h2>
                            <p style='text-align:center;color:#666;'>" + ex.Message + @"</p>
                            <p style='text-align:center;margin-top:20px;'>请尝试使用""导出Excel""按钮导出报告后查看</p></body></html>";
                        
                        string tempErrorFile = Path.Combine(Path.GetTempPath(), $"error_preview_{Guid.NewGuid()}.html");
                        File.WriteAllText(tempErrorFile, errorHtml);
                        WebPreview.Navigate(new Uri(tempErrorFile));
                    }
                    catch
                    {
                        // 如果连错误HTML都无法显示，直接显示消息框
                        MessageBox.Show($"显示预览时出错: {ex.Message}", "预览错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
        }
        
        /// <summary>
        /// 增强HTML，添加额外的图片支持和样式
        /// </summary>
        private string EnhanceHtmlWithImageSupport(string originalHtml)
        {
            // 如果HTML为空，直接返回
            if (string.IsNullOrEmpty(originalHtml))
                return originalHtml;
                
            try
            {
                // 添加特定的样式和脚本增强HTML
                if (!originalHtml.Contains("<head>"))
                    return originalHtml;
                
                // 找到</head>的位置
                int headEndIndex = originalHtml.IndexOf("</head>");
                if (headEndIndex <= 0)
                    return originalHtml;
                    
                // 构建额外的样式内容
                StringBuilder extraStyles = new StringBuilder();
                extraStyles.AppendLine("<style>");
                extraStyles.AppendLine("@media print {");
                extraStyles.AppendLine("  body { margin: 0; padding: 0; }");
                extraStyles.AppendLine("  table { width: 100%; }");
                extraStyles.AppendLine("}");
                extraStyles.AppendLine(".report-background-image {");
                extraStyles.AppendLine("  position: fixed;");
                extraStyles.AppendLine("  top: 0;");
                extraStyles.AppendLine("  left: 0;");
                extraStyles.AppendLine("  width: 100%;");
                extraStyles.AppendLine("  height: 100%;");
                extraStyles.AppendLine("  z-index: -1;");
                extraStyles.AppendLine("  opacity: 0.2;");
                extraStyles.AppendLine("  pointer-events: none;");
                extraStyles.AppendLine("}");
                extraStyles.AppendLine("</style>");
                
                // 将样式添加到head中
                string enhancedHtml = originalHtml.Insert(headEndIndex, extraStyles.ToString());
                
                // 如果不包含背景图片提示，添加一个背景图片示例
                if (!enhancedHtml.Contains("<div class='report-content'>") && !enhancedHtml.Contains("report-background"))
                {
                    int bodyStartIndex = enhancedHtml.IndexOf("<body>");
                    if (bodyStartIndex > 0)
                    {
                        // 在body开始处添加一个提示，Excel版本中可能包含更好的格式和图片
                        string bodyAddition = "\n<div style='text-align:right;font-size:12px;color:#666;margin:5px;'>Excel版本包含更完整的格式和图片</div>\n";
                        enhancedHtml = enhancedHtml.Insert(bodyStartIndex + 6, bodyAddition);
                    }
                }
                
                return enhancedHtml;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"增强HTML时出错: {ex.Message}");
                return originalHtml; // 出错时返回原始HTML
            }
        }

        /// <summary>
        /// 导出Excel按钮点击事件
        /// </summary>
        private void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(_excelFilePath) || !File.Exists(_excelFilePath))
                {
                    MessageBox.Show("Excel文件不存在，无法导出。", "导出错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }
                
                // 创建另存为对话框
                SaveFileDialog saveDialog = new SaveFileDialog
                {
                    Filter = "Excel文件 (*.xlsx)|*.xlsx",
                    Title = "导出Excel报告",
                    FileName = Path.GetFileName(_excelFilePath)
                };
                
                // 显示对话框并获取结果
                if (saveDialog.ShowDialog() == true)
                {
                    // 复制Excel文件到选择的路径
                    File.Copy(_excelFilePath, saveDialog.FileName, true);
                    
                    // 询问是否打开文件
                    if (MessageBox.Show("报告已导出，是否打开？", "导出成功", MessageBoxButton.YesNo, MessageBoxImage.Information) == MessageBoxResult.Yes)
                    {
                        Process.Start(new ProcessStartInfo(saveDialog.FileName) { UseShellExecute = true });
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"导出过程中发生错误：{ex.Message}", "导出错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 打印按钮点击事件
        /// </summary>
        private void BtnPrint_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // 使用WebBrowser的打印功能
                dynamic activeX = WebPreview.GetType().InvokeMember("ActiveXInstance",
                    System.Reflection.BindingFlags.GetProperty | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic,
                    null, WebPreview, new object[] { });
                
                if (activeX != null)
                {
                    activeX.ExecWB(6, 1); // 6表示打印，1表示显示打印对话框
                }
                else
                {
                    MessageBox.Show("无法获取打印接口，请尝试导出后打印。", "打印错误", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"打印过程中发生错误：{ex.Message}", "打印错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 上一页按钮点击事件
        /// </summary>
        private void BtnPrevPage_Click(object sender, RoutedEventArgs e)
        {
            if (_currentPage > 0)
            {
                _currentPage--;
                ShowCurrentPage();
            }
        }

        /// <summary>
        /// 下一页按钮点击事件
        /// </summary>
        private void BtnNextPage_Click(object sender, RoutedEventArgs e)
        {
            if (_currentPage < _htmlReports.Count - 1)
            {
                _currentPage++;
                ShowCurrentPage();
            }
        }
    }
} 