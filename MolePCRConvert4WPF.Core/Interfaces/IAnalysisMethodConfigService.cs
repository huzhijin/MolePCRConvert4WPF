using MolePCRConvert4WPF.Core.Models;
using System.Collections.ObjectModel;
using System.Threading.Tasks;

namespace MolePCRConvert4WPF.Core.Interfaces
{
    /// <summary>
    /// Service interface for managing analysis method configurations stored in Excel files.
    /// </summary>
    public interface IAnalysisMethodConfigService
    {
        /// <summary>
        /// Reads the analysis method configuration from the specified Excel file.
        /// </summary>
        /// <param name="filePath">The full path to the Excel file.</param>
        /// <returns>A collection of analysis method rules.</returns>
        /// <exception cref="FileNotFoundException">Thrown if the file does not exist.</exception>
        /// <exception cref="Exception">Thrown if there is an error reading the file.</exception>
        Task<ObservableCollection<AnalysisMethodRule>> LoadConfigurationAsync(string filePath);

        /// <summary>
        /// Saves the provided analysis method configuration to the specified Excel file.
        /// </summary>
        /// <param name="filePath">The full path to the Excel file.</param>
        /// <param name="configuration">The collection of analysis method rules to save.</param>
        /// <returns>Task representing the asynchronous operation.</returns>
        /// <exception cref="ArgumentNullException">Thrown if configuration is null.</exception>
        /// <exception cref="Exception">Thrown if there is an error writing the file.</exception>
        Task SaveConfigurationAsync(string filePath, ObservableCollection<AnalysisMethodRule> configuration);

        /// <summary>
        /// Creates a new Excel file with a default template for analysis method configuration.
        /// </summary>
        /// <param name="filePath">The full path where the new Excel file should be created.</param>
        /// <returns>Task representing the asynchronous operation.</returns>
        /// <exception cref="ArgumentException">Thrown if the file path is invalid or the file already exists and overwrite is not allowed (implementation detail).</exception>
        /// <exception cref="Exception">Thrown if there is an error creating the file.</exception>
        Task CreateNewConfigurationFileAsync(string filePath);
    }
} 