using MolePCRConvert4WPF.Core.Models;
using System.Collections.ObjectModel;

namespace MolePCRConvert4WPF.Core.Services
{
    /// <summary>
    /// Interface for managing shared application state.
    /// </summary>
    public interface IAppStateService
    {
        /// <summary>
        /// Gets or sets the currently loaded plate data.
        /// </summary>
        Plate? CurrentPlate { get; set; }

        /// <summary>
        /// Gets or sets the full path to the analysis method configuration file 
        /// currently being used or selected.
        /// </summary>
        string? CurrentAnalysisMethodPath { get; set; }

        /// <summary>
        /// Gets or sets the list of available analysis method files 
        /// discovered in the configured folder.
        /// </summary>
        ObservableCollection<FileDisplayInfo>? AnalysisMethodFiles { get; set; }

        // Add other shared state properties as needed
    }
} 