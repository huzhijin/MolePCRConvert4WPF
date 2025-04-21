using MaterialDesignThemes.Wpf;
using Microsoft.Extensions.DependencyInjection;
using MolePCRConvert4WPF.Core.Services;
using System;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace MolePCRConvert4WPF.App.Services
{
    /// <summary>
    /// Implementation of IDialogService using MaterialDesignThemes DialogHost.
    /// Requires a DialogHost control in the main XAML structure with a specified Identifier.
    /// </summary>
    public class DialogService : IDialogService
    {
        private readonly IServiceProvider _serviceProvider;
        private const string DefaultDialogHostIdentifier = "MainDialogHost"; // Match this with your XAML

        public DialogService(IServiceProvider serviceProvider)
        {
            _serviceProvider = serviceProvider;
        }

        public async Task<object?> ShowDialogAsync(object viewModel, string? dialogIdentifier = null)
        {
            var view = FindViewForViewModel(viewModel.GetType());
            if (view == null)
            {
                throw new InvalidOperationException($"Could not find View for ViewModel type {viewModel.GetType().FullName}");
            }

            view.DataContext = viewModel;

            string hostId = dialogIdentifier ?? DefaultDialogHostIdentifier;

            // Ensure this runs on the UI thread if called from a background thread
            object? result = null;
            await Application.Current.Dispatcher.InvokeAsync(async () =>
            {
                 result = await DialogHost.Show(view, hostId);
                 // The result passed back from DialogHost.CloseDialogCommand's CommandParameter
            }); 
            
            return result;
        }

        /// <summary>
        /// Finds the corresponding View type for a given ViewModel type based on naming convention.
        /// Assumes View is in '.Views' namespace and ViewModel is in '.ViewModels', 
        /// and View name is ViewModel name without "Model" suffix.
        /// Example: SampleNamingViewModel -> SampleNamingView
        /// Adjust this logic based on your project's structure and naming conventions.
        /// </summary>
        private UserControl? FindViewForViewModel(Type viewModelType)
        {
            // Basic naming convention logic (adjust as needed)
            string? viewTypeName = viewModelType.FullName?
                .Replace(".ViewModels.", ".Views.")
                .Replace("ViewModel", ""); 

            if (string.IsNullOrEmpty(viewTypeName))
            {
                return null;
            }

            Type? viewType = viewModelType.Assembly.GetType(viewTypeName);

            if (viewType == null)
            {
                 // Try removing "Dialog" from name if using Dialogs subfolder for VMs
                 viewTypeName = viewTypeName.Replace("Dialogs.Dialog", "Dialogs");
                 viewType = viewModelType.Assembly.GetType(viewTypeName);
                 
                 if (viewType == null) return null;
            }

            // Use DI container to create the view instance if it has dependencies,
            // otherwise, Activator.CreateInstance is simpler.
            // Using Activator here for simplicity, assuming views don't have complex constructor injection.
             var viewInstance = Activator.CreateInstance(viewType) as UserControl;
             // Or using DI:
             // var viewInstance = _serviceProvider.GetService(viewType) as UserControl;

             return viewInstance;
        }
    }
} 