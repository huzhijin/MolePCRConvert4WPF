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
using MolePCRConvert4WPF.Core.Interfaces;

// Correct Namespace
namespace MolePCRConvert4WPF.App.Services 
{
    /// <summary>
    /// 导航服务实现
    /// </summary>
    public class NavigationService : INavigationService
    {
        private readonly IServiceProvider _serviceProvider;
        private readonly ILogger<NavigationService> _logger;
        private ContentControl? _mainContentControl; // Reference to the content area
        private readonly Dictionary<Type, Type> _viewModelToViewMap = new Dictionary<Type, Type>();
        private readonly Dictionary<Type, Action> _navigationActions = new();

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

        /// <summary>
        /// 注册视图模型对应的导航操作
        /// </summary>
        /// <typeparam name="TViewModel">视图模型类型</typeparam>
        /// <param name="navigationAction">导航操作</param>
        public void RegisterNavigationAction<TViewModel>(Action navigationAction) where TViewModel : class
        {
            _navigationActions[typeof(TViewModel)] = navigationAction;
        }

        /// <summary>
        /// 导航到指定视图模型
        /// </summary>
        public void NavigateTo<TViewModel>() where TViewModel : class
        {
            var viewModelType = typeof(TViewModel);
            if (_navigationActions.TryGetValue(viewModelType, out var action))
            {
                action?.Invoke();
            }
            else
            {
                throw new InvalidOperationException($"未注册视图模型 {viewModelType.Name} 的导航操作");
            }
        }
        
        /// <summary>
        /// 获取指定类型的视图模型实例
        /// </summary>
        public TViewModel? GetViewModel<TViewModel>() where TViewModel : class
        {
            return _serviceProvider.GetService<TViewModel>();
        }
    }
} 