using MolePCRConvert4WPF.Core.Models;
using MolePCRConvert4WPF.Core.Services;
using System.Collections.ObjectModel;

namespace MolePCRConvert4WPF.Infrastructure.Services
{
    /// <summary>
    /// Manages shared application state.
    /// Implements IAppStateService.
    /// </summary>
    public class AppStateService : IAppStateService
    {
        /// <summary>
        /// Gets or sets the currently loaded plate data.
        /// </summary>
        public Plate? CurrentPlate { get; set; }

        /// <summary>
        /// Gets or sets the full path to the analysis method configuration file
        /// currently being used or selected.
        /// </summary>
        public string? CurrentAnalysisMethodPath { get; set; }

        /// <summary>
        /// Gets or sets the list of available analysis method files 
        /// discovered in the configured folder.
        /// </summary>
        public ObservableCollection<FileDisplayInfo>? AnalysisMethodFiles { get; set; }

        // Constructor or other methods can be added if needed
        public AppStateService()
        {
            // Initialize default values if necessary
            CurrentPlate = null; // Or new Plate(); depending on requirements
            CurrentAnalysisMethodPath = null;
            AnalysisMethodFiles = new ObservableCollection<FileDisplayInfo>(); // Initialize empty collection
        }
    }
} 