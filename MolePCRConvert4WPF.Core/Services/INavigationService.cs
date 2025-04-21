using System;

namespace MolePCRConvert4WPF.Core.Services
{
    /// <summary>
    /// Defines methods for navigating between views/viewmodels within the application.
    /// </summary>
    public interface INavigationService
    {
        /// <summary>
        /// Navigates to the view associated with the specified ViewModel type.
        /// </summary>
        /// <typeparam name="TViewModel">The type of the ViewModel to navigate to.</typeparam>
        void NavigateTo<TViewModel>() where TViewModel : class; // Using class constraint for ViewModels

        // Optional: Add other navigation methods if needed, e.g.:
        // void GoBack();
        // bool CanGoBack { get; }
        // event EventHandler<string> Navigated; // Event raised after navigation
    }
} 