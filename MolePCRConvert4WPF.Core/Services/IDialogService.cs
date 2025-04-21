using System.Threading.Tasks;

namespace MolePCRConvert4WPF.Core.Services
{
    /// <summary>
    /// Interface for showing dialogs.
    /// </summary>
    public interface IDialogService
    {
        /// <summary>
        /// Shows a dialog associated with the provided ViewModel.
        /// </summary>
        /// <param name="viewModel">The ViewModel for the dialog content.</param>
        /// <param name="dialogIdentifier">Optional identifier for targeting a specific DialogHost instance.</param>
        /// <returns>The result from the dialog. Typically the ViewModel itself if confirmed, null/false/specific object if cancelled.</returns>
        Task<object?> ShowDialogAsync(object viewModel, string? dialogIdentifier = null);
    }
} 