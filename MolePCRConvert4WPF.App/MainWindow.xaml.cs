using System;
using System.Windows;
using System.Windows.Controls;
using MahApps.Metro.Controls;
using System.Diagnostics;
using MolePCRConvert4WPF.App.Views.DataInput;
using MolePCRConvert4WPF.App.Views.SampleAnalysis;
using MolePCRConvert4WPF.App.Views.AnalysisMethodConfig;
using MolePCRConvert4WPF.App.ViewModels;
using Microsoft.Extensions.DependencyInjection;
using MolePCRConvert4WPF.App.Views;
using MolePCRConvert4WPF.Core.Services;
using MolePCRConvert4WPF.App.Services;

namespace MolePCRConvert4WPF.App
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : MetroWindow
    {
        private readonly INavigationService _navigationService;
        
        public MainWindow()
        {
            try
            {
                Debug.WriteLine("开始初始化MainWindow...");
                
                // 初始化组件
                InitializeComponent();
                
                Debug.WriteLine("InitializeComponent完成");
                
                // 确保MainContent已正确初始化
                if (MainContent == null)
                {
                    Debug.WriteLine("警告：MainContent为null，尝试手动初始化");
                    MainContent = new ContentControl();
                    // 获取Grid并添加MainContent
                    var mainGrid = this.Content as Grid;
                    if (mainGrid != null && mainGrid.Children.Count > 1)
                    {
                        var contentGrid = mainGrid.Children[1] as Grid;
                        if (contentGrid != null && contentGrid.Children.Count > 1)
                        {
                            contentGrid.Children[1] = MainContent;
                            Debug.WriteLine("已手动初始化MainContent");
                        }
                    }
                }
                
                // 设置标题
                Title = "PCR结果分析与报告生成系统";
                
                Debug.WriteLine("MainWindow初始化完成");
                
                // Get NavigationService from App static property
                _navigationService = App.NavigationService 
                    ?? throw new InvalidOperationException("NavigationService not initialized in App.xaml.cs");
                
                // Initialize NavigationService with the ContentControl
                if (_navigationService is MolePCRConvert4WPF.App.Services.NavigationService concreteNavService)
                {
                    concreteNavService.RegisterContentControl(MainContent); // Pass the ContentControl - Changed from Initialize
                    
                    // 注册各个ViewModel的导航操作
                    concreteNavService.RegisterNavigationAction<DataInputViewModel>(() => {
                        // 使用MainContent的依赖注入创建视图
                        var vm = concreteNavService.GetViewModel<DataInputViewModel>();
                        var view = new Views.DataInput.DataInputView { DataContext = vm };
                        MainContent.Content = view;
                        Debug.WriteLine("导航到DataInputView完成");
                    });
                    
                    concreteNavService.RegisterNavigationAction<SampleAnalysisViewModel>(() => {
                        var vm = concreteNavService.GetViewModel<SampleAnalysisViewModel>();
                        var view = new Views.SampleAnalysis.SampleAnalysisView { DataContext = vm };
                        MainContent.Content = view;
                        Debug.WriteLine("导航到SampleAnalysisView完成");
                    });
                    
                    concreteNavService.RegisterNavigationAction<ExcelAnalysisMethodConfigViewModel>(() => {
                        var vm = concreteNavService.GetViewModel<ExcelAnalysisMethodConfigViewModel>();
                        var view = new Views.AnalysisMethodConfig.ExcelAnalysisMethodConfigView { DataContext = vm };
                        MainContent.Content = view;
                        Debug.WriteLine("导航到ExcelAnalysisMethodConfigView完成");
                    });
                    
                    concreteNavService.RegisterNavigationAction<PCRResultAnalysisViewModel>(() => {
                        var vm = concreteNavService.GetViewModel<PCRResultAnalysisViewModel>();
                        var view = new Views.ResultAnalysis.PCRResultAnalysisView { DataContext = vm };
                        MainContent.Content = view;
                        Debug.WriteLine("导航到PCRResultAnalysisView完成");
                    });
                    
                    concreteNavService.RegisterNavigationAction<ReportTemplateConfigViewModel>(() => {
                        // 报告模板配置视图
                        var vm = concreteNavService.GetViewModel<ReportTemplateConfigViewModel>();
                        var view = new Views.ReportTemplateConfigView { DataContext = vm };
                        MainContent.Content = view;
                        Debug.WriteLine("导航到ReportTemplateConfigView完成");
                    });
                    
                    concreteNavService.RegisterNavigationAction<ReportTemplateDesignerViewModel>(() => {
                        // 报告模板设计器视图
                        var vm = concreteNavService.GetViewModel<ReportTemplateDesignerViewModel>();
                        var view = new Views.ReportTemplateDesignerView { DataContext = vm };
                        MainContent.Content = view;
                        Debug.WriteLine("导航到ReportTemplateDesignerView完成");
                    });
                }
                else
                {
                    Debug.WriteLine("WARNING: Resolved INavigationService is not the expected concrete type (App.Services.NavigationService) for initialization.");
                }
                
                // Attach event handlers and perform initial navigation AFTER component initialization and service setup
                this.Loaded += (s, e) =>
                {
                    try
                    {
                        // Attach the event handler here, AFTER the listbox is loaded
                        NavListBox.SelectionChanged += NavListBox_SelectionChanged;
                        Debug.WriteLine("NavListBox_SelectionChanged handler attached.");
                        
                        Debug.WriteLine("MainWindow loaded, attempting initial navigation...");
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            ListBoxItem? targetItem = null;
                            // 找到数据输入菜单
                            foreach (var menuItem in NavListBox.Items)
                            {
                                if (menuItem is ListBoxItem listItem && listItem.Tag?.ToString() == "DataInput")
                                {
                                    targetItem = listItem;
                                    break;
                                }
                            }

                            if (targetItem != null)
                            {                            
                                // Set SelectedItem which might trigger SelectionChanged now
                                NavListBox.SelectedItem = targetItem;
                                Debug.WriteLine("Initial SelectedItem set to DataInput.");
                                
                                // If SelectionChanged didn't fire (e.g., already selected), navigate manually.
                                // We add a check to see if content is already loaded by the event handler.
                                if (MainContent.Content == null || !(MainContent.Content is Views.DataInput.DataInputView))
                                {
                                     Debug.WriteLine("Manually calling NavigateToItem for initial load.");
                                     NavigateToItem(targetItem); 
                                }
                                else
                                {
                                     Debug.WriteLine("MainContent already populated, skipping manual navigation call.");
                                }
                            }
                            else 
                            {
                                Debug.WriteLine("WARNING：未在NavListBox中找到Tag为'DataInput'的项");
                                // Load welcome view manually if DataInput not found
                                MainContent.Content = CreateWelcomeView();
                            }
                        }), System.Windows.Threading.DispatcherPriority.Loaded);
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"Error in MainWindow Loaded event: {ex.Message}");
                         MessageBox.Show($"窗口加载时出错: {ex.Message}", "加载错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                };
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"MainWindow构造函数出现未处理异常: {ex.Message}\n{ex.StackTrace}");
                MessageBox.Show($"初始化主窗口时出错：{ex.Message}\n{ex.StackTrace}", "严重错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void NavListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Add null check for navigation service just in case
            if (_navigationService == null)
            {
                 Debug.WriteLine("ERROR in NavListBox_SelectionChanged: _navigationService is null.");
                 MessageBox.Show("导航服务尚未初始化！", "严重错误", MessageBoxButton.OK, MessageBoxImage.Error);
                 return;
            }
            
            if (e.AddedItems.Count > 0 && e.AddedItems[0] is ListBoxItem selectedItem)
            {
                 NavigateToItem(selectedItem);
            }
            // Removed fallback as initial navigation is handled in Loaded
            // else if (NavListBox.SelectedItem is ListBoxItem fallbackItem) ... 
        }

        // Modify NavigateToItem to use the INavigationService
        private void NavigateToItem(ListBoxItem item)
        {
            // Add null check for navigation service just in case
            if (_navigationService == null)
            {
                 Debug.WriteLine("ERROR in NavigateToItem: _navigationService is null.");
                 MessageBox.Show("导航服务尚未初始化！", "严重错误", MessageBoxButton.OK, MessageBoxImage.Error);
                 return;
            }
            
            if (item?.Tag is string tag)
            {
                try
                {
                     Debug.WriteLine($"Navigating via NavigationService for tag: {tag}");
                     // Map tag to ViewModel type and call NavigationService
                     switch (tag)
                     {
                         case "DataInput":
                             _navigationService.NavigateTo<DataInputViewModel>();
                             break;
                         case "SampleAnalysis":
                             _navigationService.NavigateTo<SampleAnalysisViewModel>();
                             break;
                         case "ExcelAnalysisMethod":
                             _navigationService.NavigateTo<ExcelAnalysisMethodConfigViewModel>();
                             break;
                         case "PCRResultAnalysis":
                             _navigationService.NavigateTo<PCRResultAnalysisViewModel>();
                             break;
                         case "ReportTemplateConfig":
                             _navigationService.NavigateTo<ReportTemplateConfigViewModel>();
                             break;
                         case "ReportTemplateDesigner":
                             _navigationService.NavigateTo<ReportTemplateDesignerViewModel>();
                             break;
                         default:
                             Debug.WriteLine($"Unknown navigation tag: {tag}. Displaying welcome view.");
                             // Optionally navigate to a default/welcome view via service if needed
                             // _navigationService.NavigateTo<WelcomeViewModel>(); 
                             // Or clear content manually if service doesn't handle default
                             MainContent.Content = CreateWelcomeView(); 
                             break;
                     }
                 }
                 catch (Exception ex)
                 {
                     Debug.WriteLine($"Error during NavigateToItem call: {ex.Message}\n{ex.StackTrace}");
                     MessageBox.Show($"导航处理时发生错误: {ex.Message}", "导航错误", MessageBoxButton.OK, MessageBoxImage.Error);
                     MainContent.Content = CreateWelcomeView(); // Fallback
                 }
            }
            else
            {
                 Debug.WriteLine("NavigateToItem called with invalid item or tag.");
            }
        }

        private void ShowNotImplementedMessage(string feature)
        {
            var grid = new Grid();
            var panel = new StackPanel
            {
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center
            };
            
            panel.Children.Add(new TextBlock
            {
                Text = $"{feature}功能正在开发中...",
                FontSize = 24,
                TextAlignment = TextAlignment.Center,
                Margin = new Thickness(0, 0, 0, 20)
            });
            
            panel.Children.Add(new TextBlock
            {
                Text = "请在将来的版本中期待此功能",
                FontSize = 16,
                TextAlignment = TextAlignment.Center
            });
            
            grid.Children.Add(panel);
            MainContent.Content = grid;
        }

        private UIElement CreateWelcomeView()
        {
            var grid = new Grid();
            var panel = new StackPanel
            {
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center
            };
            
            panel.Children.Add(new TextBlock
            {
                Text = "欢迎使用PCR结果分析与报告生成系统",
                FontSize = 24,
                TextAlignment = TextAlignment.Center,
                Margin = new Thickness(0, 0, 0, 20)
            });
            
            panel.Children.Add(new TextBlock
            {
                Text = "请从左侧菜单选择操作",
                FontSize = 16,
                TextAlignment = TextAlignment.Center
            });
            
            grid.Children.Add(panel);
            return grid;
        }
    }
} 