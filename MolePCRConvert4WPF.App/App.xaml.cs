using System;
using System.Collections.Generic;
//using System.Configuration;
using System.IO;
using System.Linq;
using System.Windows;
using Microsoft.Extensions.DependencyInjection;
using MolePCRConvert4WPF.Core.Interfaces;
using MolePCRConvert4WPF.Core.Models;
using MolePCRConvert4WPF.Core.Models.PanelRules;
using System.Text.Json;
using OfficeOpenXml;
using MolePCRConvert4WPF.Core.Enums;
using LicenseContext = OfficeOpenXml.LicenseContext;
using MolePCRConvert4WPF.Core.Services;
using MolePCRConvert4WPF.Infrastructure.FileHandlers;
using MolePCRConvert4WPF.App.ViewModels;
using Microsoft.Extensions.Logging;
using MolePCRConvert4WPF.Infrastructure.Services;
using MolePCRConvert4WPF.App.Services;
using System.Diagnostics;

namespace MolePCRConvert4WPF.App
{
    /// <summary>
    /// App.xaml 的交互逻辑
    /// </summary>
    public partial class App : Application
    {
        // 配置文件路径
        private readonly string _panelRulesFilePath = "PanelRules.json";
        private readonly string _appSettingsFilePath = "AppSettings.json";
        
        // 应用程序设置
        private Dictionary<string, object>? _appSettings;
        
        // 服务提供者
        public static IServiceProvider? ServiceProvider { get; private set; }

        // 当前板数据（跨视图共享）
        public Plate? CurrentPlate { get; set; }
        
        // 规则引擎
        public RuleEngine? RuleEngine { get; private set; }

        // Static property to hold the navigation service instance after it's configured
        public static INavigationService? NavigationService { get; private set; }

        public App()
        {
            _appSettings = new Dictionary<string, object>();
            CurrentPlate = new Plate();
            RuleEngine = new RuleEngine();
        }

        /// <summary>
        /// 设置EPPlus许可证上下文
        /// </summary>
        public static bool SetupEPPlusLicense()
        {
            try
            {
                // 设置EPPlus许可证上下文为非商业用途
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                
                System.Diagnostics.Debug.WriteLine("EPPlus许可证上下文已成功设置为NonCommercial");
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"设置EPPlus许可证上下文时出错: {ex.Message}");
                return false;
            }
        }

        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);
            
            // 设置应用程序异常处理
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
            Application.Current.DispatcherUnhandledException += Current_DispatcherUnhandledException;
            
            System.Diagnostics.Debug.WriteLine("===== 应用程序启动 =====");
            
            try
            {
                System.Diagnostics.Debug.WriteLine("正在初始化(或加载)依赖于服务的对象...");
                InitializeAppSettings();
                InitializePanelRules();
                System.Diagnostics.Debug.WriteLine("依赖于服务的对象初始化完成");
            }
            catch (Exception objEx)
            {
                System.Diagnostics.Debug.WriteLine($"初始化基本对象时出错: {objEx.Message}");
                MessageBox.Show($"初始化基本对象时出错: {objEx.Message}\n应用程序可能无法正常工作。", 
                    "初始化错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            
            try
            {
                System.Diagnostics.Debug.WriteLine("正在配置服务...");
                ConfigureServices();
                System.Diagnostics.Debug.WriteLine("服务配置完成");
                
                // Get the NavigationService instance AFTER the container is built
                NavigationService = ServiceProvider?.GetService<INavigationService>();
                if (NavigationService == null) 
                {
                    Debug.WriteLine("CRITICAL ERROR: Failed to resolve INavigationService after container build.");
                    // Handle critical error, maybe shutdown
                }
            }
            catch (Exception servicesEx)
            {
                System.Diagnostics.Debug.WriteLine($"配置服务时出错: {servicesEx.Message}");
                MessageBox.Show($"配置服务时出错: {servicesEx.Message}\n应用程序将继续使用有限功能启动。", 
                    "初始化警告", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            
            // 显示主窗口
            try
            {
                System.Diagnostics.Debug.WriteLine("正在创建并显示主窗口...");
                
                if (Current.MainWindow != null)
                {
                    System.Diagnostics.Debug.WriteLine($"警告：当前已存在MainWindow实例，将使用现有实例");
                    Current.MainWindow.Visibility = Visibility.Visible;
                    Current.MainWindow.Activate();
                    System.Diagnostics.Debug.WriteLine("使用已有MainWindow实例并激活");
                }
                else
                {
                    // 使用Dispatcher确保在UI线程上创建窗口
                    Dispatcher.Invoke(() =>
                    {
                        try
                        {
                            MainWindow mainWindow = new MainWindow();
                            Current.MainWindow = mainWindow;
                            mainWindow.Show();
                            System.Diagnostics.Debug.WriteLine("主窗口已成功创建并显示");
                        }
                        catch (Exception windowEx)
                        {
                            System.Diagnostics.Debug.WriteLine($"创建或显示主窗口时出错: {windowEx.Message}");
                            MessageBox.Show($"创建或显示主窗口时出错: {windowEx.Message}", 
                                "严重错误", MessageBoxButton.OK, MessageBoxImage.Error);
                            Current.Shutdown();
                        }
                    });
                }
            }
            catch (Exception winEx)
            {
                System.Diagnostics.Debug.WriteLine($"显示主窗口过程中出错: {winEx.Message}");
                MessageBox.Show($"显示主窗口过程中出错: {winEx.Message}", 
                    "严重错误", MessageBoxButton.OK, MessageBoxImage.Error);
                Current.Shutdown();
            }
        }

        private void InitializePanelRules()
        {
            if (ServiceProvider == null)
            {
                System.Diagnostics.Debug.WriteLine("错误: ServiceProvider 未初始化，无法加载 Panel 规则。");
                RuleEngine!.Configuration = CreateDefaultPanelRuleConfig();
                return;
            }

            try
            {
                var ruleRepository = ServiceProvider.GetService<IPanelRuleRepository>();
                if (ruleRepository != null)
                {
                    string defaultPanelName = "默认分析规则";
                    var config = ruleRepository.GetPanelRuleConfiguration(defaultPanelName);
                    
                    if (config != null)
                    { 
                        RuleEngine!.Configuration = config;
                        System.Diagnostics.Debug.WriteLine($"成功从仓库加载规则配置: {config.Name}");
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine($"仓库中未找到规则配置 '{defaultPanelName}'，将使用默认配置。");
                        RuleEngine!.Configuration = CreateDefaultPanelRuleConfig();
                    }
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("错误: 未能从 ServiceProvider 获取 IPanelRuleRepository。将使用默认配置。");
                    RuleEngine!.Configuration = CreateDefaultPanelRuleConfig();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"初始化面板规则时出错: {ex.Message}");
                RuleEngine!.Configuration = CreateDefaultPanelRuleConfig();
            }
        }

        private PanelRuleConfiguration CreateDefaultPanelRuleConfig()
        {
            return new PanelRuleConfiguration
            {
                Name = "默认分析规则",
                Description = "系统默认分析规则配置",
                Version = "1.0",
                RuleGroups = new List<RuleGroup>
                {
                    new RuleGroup
                    {
                        Name = "默认规则组",
                        Priority = 1,
                        Rules = new List<Rule>
                        {
                            new Rule
                            {
                                Name = "默认阳性规则",
                                Condition = "Channel == 'FAM' && CtValue >= 10 && CtValue <= 35",
                                Action = "Result = Positive"
                            },
                            new Rule
                            {
                                Name = "默认阴性规则",
                                Condition = "Channel == 'FAM' && (CtValue < 10 || CtValue > 35 || CtValue == null)",
                                Action = "Result = Negative"
                            }
                        }
                    }
                },
                Channels = new List<ChannelConfiguration>
                {
                    new ChannelConfiguration
                    {
                        Name = "FAM",
                        Target = "Target1",
                        MinPositiveCt = 10,
                        MaxPositiveCt = 35
                    },
                    new ChannelConfiguration
                    {
                        Name = "VIC",
                        Target = "Target2",
                        MinPositiveCt = 10,
                        MaxPositiveCt = 35
                    }
                }
            };
        }

        private void ConfigureServices()
        {
            System.Diagnostics.Debug.WriteLine("ConfigureServices: 开始配置服务");
            var services = new ServiceCollection();

            // 添加日志
            services.AddLogging(builder =>
            {
                builder.AddDebug(); // 将日志输出到调试输出窗口
                // 可选：添加其他日志提供程序，例如文件或控制台
            });

            // 注册核心服务
            System.Diagnostics.Debug.WriteLine("ConfigureServices: 注册核心服务");
            services.AddSingleton<IPanelRuleRepository>(provider => 
                new PanelRuleRepository(_panelRulesFilePath));
            services.AddTransient<ISampleNamingService, SampleNamingService>();
            services.AddTransient<IReportService, ExcelReportGenerator>();
            services.AddTransient<IReportTemplateDesignerService, ReoGridReportTemplateDesignerService>();
            services.AddTransient<IFileHandler, ExcelFileHandler>();

            // 注册视图模型
            System.Diagnostics.Debug.WriteLine("ConfigureServices: 注册视图模型");
            services.AddTransient<DataInputViewModel>();
            services.AddTransient<SampleAnalysisViewModel>();
            services.AddTransient<ExcelAnalysisMethodConfigViewModel>();
            services.AddTransient<PCRResultAnalysisViewModel>();
            services.AddTransient<ReportTemplateConfigViewModel>();
            services.AddTransient<ReportTemplateDesignerViewModel>();
            // 注册 ShellViewModel (假设它存在且需要被注入)
            // 如果 ShellViewModel 不直接通过 DI 创建，则不需要下面这行
            // services.AddTransient<ShellViewModel>(); 

            // 注册应用级服务（如果需要）
            // services.AddSingleton<ISettingsService, SettingsService>();
            // 注册新添加的服务
            System.Diagnostics.Debug.WriteLine("ConfigureServices: 注册应用状态和导航服务");
            services.AddSingleton<IAppStateService, AppStateService>();
            services.AddSingleton<INavigationService, NavigationService>();
            services.AddSingleton<IUserSettingsService, UserSettingsService>();
            
            // 注册 PCRResultAnalysisViewModel 依赖的服务
            System.Diagnostics.Debug.WriteLine("ConfigureServices: 注册分析方法配置和PCR分析服务");
            services.AddTransient<IAnalysisMethodConfigService, NpoiAnalysisMethodConfigService>(); 

            // 先注册具体的PCRAnalysisService，确保工厂可以获取具体类型实例
            services.AddSingleton<PCRAnalysisService>();

            // 注册SLAN专用的PCR分析服务
            services.AddSingleton<SLANPCRAnalysisService>();

            // 注册PCR分析服务工厂
            services.AddSingleton<IPCRAnalysisServiceFactory, PCRAnalysisServiceFactory>();

            // 注册PCR分析服务接口
            services.AddSingleton<IPCRAnalysisService>(provider => 
                provider.GetRequiredService<IPCRAnalysisServiceFactory>().GetAnalysisService(InstrumentType.Unknown));

            System.Diagnostics.Debug.WriteLine("ConfigureServices: 构建服务提供程序");
            ServiceProvider = services.BuildServiceProvider();
            System.Diagnostics.Debug.WriteLine("ConfigureServices: 服务配置完成");
        }

        protected override void OnExit(ExitEventArgs e)
        {
            // 释放服务
            if (ServiceProvider is IDisposable disposable)
            {
                disposable.Dispose();
            }
            
            base.OnExit(e);
        }

        private void InitializeAppSettings()
        {
            try
            {
                if (File.Exists(_appSettingsFilePath))
                {
                    string json = File.ReadAllText(_appSettingsFilePath);
                    _appSettings = JsonSerializer.Deserialize<Dictionary<string, object>>(json);
                }
                else
                {
                    _appSettings = new Dictionary<string, object>();
                    File.WriteAllText(_appSettingsFilePath, JsonSerializer.Serialize(_appSettings));
                }
            }
            catch
            {
                _appSettings = new Dictionary<string, object>();
            }
        }

        private void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            HandleException(e.ExceptionObject as Exception);
        }

        private void Current_DispatcherUnhandledException(object sender, System.Windows.Threading.DispatcherUnhandledExceptionEventArgs e)
        {
            HandleException(e.Exception);
            e.Handled = true;
        }

        private void HandleException(Exception? ex)
        {
            if (ex == null) return;
            
            System.Diagnostics.Debug.WriteLine($"未处理的异常: {ex.Message}\n{ex.StackTrace}");
            
            MessageBox.Show($"发生未处理的异常: {ex.Message}\n\n请联系技术支持并提供以下信息:\n{ex.GetType().Name}\n{ex.StackTrace}",
                "应用程序错误", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }
    
    /// <summary>
    /// 规则引擎
    /// </summary>
    public class RuleEngine
    {
        /// <summary>
        /// 规则配置
        /// </summary>
        public PanelRuleConfiguration Configuration { get; set; }
        
        public RuleEngine()
        {
            Configuration = new PanelRuleConfiguration { Name="Init", RuleGroups = new(), Channels = new() };
        }
        
        // 分析PCR结果方法可以在这里添加
    }

    /// <summary>
    /// PCR分析服务工厂接口
    /// </summary>
    public interface IPCRAnalysisServiceFactory
    {
        /// <summary>
        /// 根据仪器类型获取对应的PCR分析服务
        /// </summary>
        IPCRAnalysisService GetAnalysisService(InstrumentType instrumentType);
    }
    
    /// <summary>
    /// PCR分析服务工厂实现
    /// </summary>
    public class PCRAnalysisServiceFactory : IPCRAnalysisServiceFactory
    {
        private readonly PCRAnalysisService _defaultService;
        private readonly SLANPCRAnalysisService _slanService;
        
        public PCRAnalysisServiceFactory(PCRAnalysisService defaultService, SLANPCRAnalysisService slanService)
        {
            _defaultService = defaultService;
            _slanService = slanService;
        }
        
        public IPCRAnalysisService GetAnalysisService(InstrumentType instrumentType)
        {
            // 根据仪器类型返回对应的分析服务实现
            switch (instrumentType)
            {
                case InstrumentType.SLAN96P:
                case InstrumentType.SLAN96S:
                    return _slanService;
                default:
                    return _defaultService;
            }
        }
    }
} 