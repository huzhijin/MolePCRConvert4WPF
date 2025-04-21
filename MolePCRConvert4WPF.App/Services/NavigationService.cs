using MolePCRConvert4WPF.Core.Services;
using System;
using System.Diagnostics;
using System.Windows.Controls;
using Microsoft.Extensions.DependencyInjection;
using System.Collections.Generic;
using MolePCRConvert4WPF.App.ViewModels; // Needs App reference (ok now)
using MolePCRConvert4WPF.App.Views.DataInput; // Needs App reference (ok now)
using MolePCRConvert4WPF.App.Views.SampleAnalysis; // Needs App reference (ok now)
using MolePCRConvert4WPF.App.Views.AnalysisMethodConfig; // Needs App reference (ok now)
using MolePCRConvert4WPF.App.Views.ResultAnalysis; // Added for PCRResultAnalysisView
using System.Windows;
using Microsoft.Extensions.Logging;

// Correct Namespace
namespace MolePCRConvert4WPF.App.Services 
{
    /// <summary>
    /// Implementation of INavigationService using a ContentControl.
    /// Resides in the App layer as it depends on UI elements and ViewModels.
    /// </summary>
    public class NavigationService : INavigationService
    {
        private readonly IServiceProvider _serviceProvider;
        private readonly ILogger<NavigationService> _logger;
        private ContentControl? _mainContentControl; // Reference to the content area
        private readonly Dictionary<Type, Type> _viewModelToViewMap = new Dictionary<Type, Type>();

        public NavigationService(IServiceProvider serviceProvider, ILogger<NavigationService> logger)
        { 
            _serviceProvider = serviceProvider ?? throw new ArgumentNullException(nameof(serviceProvider));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            ConfigureMappings();
            _logger.LogInformation("[NavigationService] Initialized.");
        }

        // Method to set the target ContentControl after MainWindow is created
        public void RegisterContentControl(ContentControl contentControl)
        {
            _mainContentControl = contentControl;
            _logger.LogInformation("[NavigationService] ContentControl registered: {Name}", contentControl?.Name ?? "(unnamed)");
        }

        private void ConfigureMappings()
        {
            _logger.LogInformation("[NavigationService] Configuring ViewModel to View mappings...");
            // Register ViewModel-View pairs here
            _viewModelToViewMap.Add(typeof(DataInputViewModel), typeof(DataInputView));
            _viewModelToViewMap.Add(typeof(SampleAnalysisViewModel), typeof(SampleAnalysisView));
            _viewModelToViewMap.Add(typeof(ExcelAnalysisMethodConfigViewModel), typeof(ExcelAnalysisMethodConfigView));
            _viewModelToViewMap.Add(typeof(PCRResultAnalysisViewModel), typeof(PCRResultAnalysisView));
            // Add other mappings as needed
            _logger.LogInformation("[NavigationService] ViewModel to View mappings configured.");
        }

        public void NavigateTo<TViewModel>() where TViewModel : class
        {
            // --- Add Diagnostic Checks --- 
            Debug.WriteLine($"[NavigationService.NavigateTo<{typeof(TViewModel).Name}>] Attempting navigation.");
            if (_mainContentControl == null)
            {
                Debug.WriteLine("ERROR: NavigationService cannot navigate because _mainContentControl is null.");
                MessageBox.Show("导航服务错误：内容控件引用为空！", "导航错误", MessageBoxButton.OK, MessageBoxImage.Error);
                // Trying to find the control again (less ideal, indicates an init issue)
                try { 
                    _mainContentControl = (Application.Current.MainWindow as MainWindow)?.MainContent;
                     Debug.WriteLine($"Attempted to re-acquire MainContent. Is null? {(_mainContentControl == null)}");
                } catch (Exception findEx) {
                     Debug.WriteLine($"Exception while trying to re-acquire MainContent: {findEx.Message}");
                }
                 if (_mainContentControl == null) return; // Still null, exit
            }
             if (_serviceProvider == null)
            {
                Debug.WriteLine("ERROR: NavigationService cannot navigate because _serviceProvider is null.");
                MessageBox.Show("导航服务错误：服务提供者引用为空！", "导航错误", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            Debug.WriteLine($"[NavigationService.NavigateTo] _mainContentControl: {_mainContentControl.Name ?? "(no name)"}, _serviceProvider available: {(_serviceProvider != null)}");
            // --- End Diagnostic Checks --- 

            // Original null check for _mainContentControl (redundant now but harmless)
            if (_mainContentControl == null)
            {
                Debug.WriteLine("ERROR: NavigationService cannot navigate because MainContentControl is not initialized.");
                MessageBox.Show("导航服务未正确初始化！", "导航错误", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            var viewModelType = typeof(TViewModel);
            Debug.WriteLine($"[NavigationService.NavigateTo] Looking up view for ViewModel: {viewModelType.FullName}"); // Log full name
            if (!_viewModelToViewMap.TryGetValue(viewModelType, out var viewType))
            {
                Debug.WriteLine($"ERROR: No view registered for ViewModel type: {viewModelType.Name}");
                MessageBox.Show($"找不到与 {viewModelType.Name} 关联的视图。", "导航错误", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            Debug.WriteLine($"[NavigationService.NavigateTo] Found view type: {viewType.FullName}");

            try
            {
                 Debug.WriteLine($"[NavigationService.NavigateTo] Attempting to resolve ViewModel: {viewModelType.FullName}");
                 var viewModel = _serviceProvider.GetService(viewModelType);
                 // --- Add Check for ViewModel Resolution --- 
                 if (viewModel == null)
                 {
                      Debug.WriteLine($"ERROR: _serviceProvider.GetService returned null for {viewModelType.FullName}.");
                      MessageBox.Show($"无法创建 {viewModelType.Name} 的实例。请检查服务注册及其依赖项。", "导航错误", MessageBoxButton.OK, MessageBoxImage.Error);
                      return;
                 }
                 Debug.WriteLine($"[NavigationService.NavigateTo] ViewModel resolved successfully: {viewModel.GetType().FullName}");
                 // --- End Check --- 

                 Debug.WriteLine($"[NavigationService.NavigateTo] Attempting to create View instance: {viewType.FullName}");
                 var view = Activator.CreateInstance(viewType) as UserControl;
                 // --- Add Check for View Creation --- 
                 if (view == null)
                 {
                      Debug.WriteLine($"ERROR: Activator.CreateInstance returned null or incompatible type for {viewType.FullName}.");
                      MessageBox.Show($"无法创建视图 {viewType.Name} 的实例。请检查其构造函数和基类。", "导航错误", MessageBoxButton.OK, MessageBoxImage.Error);
                      return;
                 }
                 Debug.WriteLine($"[NavigationService.NavigateTo] View instance created: {view.GetType().FullName}");
                 // --- End Check --- 
                 
                 Debug.WriteLine($"[NavigationService.NavigateTo] Setting DataContext...");
                 view.DataContext = viewModel;

                 Debug.WriteLine($"[NavigationService.NavigateTo] Dispatching UI update...");
                 _mainContentControl.Dispatcher.Invoke(() =>
                 { 
                     try
                     {
                         Debug.WriteLine($"[NavigationService.NavigateTo] Dispatcher.Invoke: Setting ContentControl Content...");
                         _mainContentControl.Content = view;
                         Debug.WriteLine($"[NavigationService.NavigateTo] Dispatcher.Invoke: Content set to {view.GetType().Name}");
                     }
                     catch(Exception dispatcherEx)
                     {
                          Debug.WriteLine($"ERROR inside Dispatcher.Invoke: {dispatcherEx.Message}\n{dispatcherEx.StackTrace}");
                          // Handle error within dispatcher if possible
                     }
                 });
            }
            catch (Exception ex)
            {
                 Debug.WriteLine($"ERROR during navigation to {viewType.Name}: {ex.Message}\n{ex.StackTrace}");
                 MessageBox.Show($"导航到 {viewType.Name} 时发生错误: {ex.Message}", "导航错误", MessageBoxButton.OK, MessageBoxImage.Error);
                 _mainContentControl.Dispatcher.Invoke(() => { _mainContentControl.Content = null; });
            }
        }
    }
} 