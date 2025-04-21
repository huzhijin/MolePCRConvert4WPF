namespace MolePCRConvert4WPF.Core.Models
{
    /// <summary>
    /// Represents information about a file, primarily for display purposes.
    /// </summary>
    public class FileDisplayInfo
    {
        /// <summary>
        /// Gets or sets the name of the file to be displayed.
        /// </summary>
        public string DisplayName { get; set; } = string.Empty;

        /// <summary>
        /// Gets or sets the full path to the file.
        /// </summary>
        public string FullPath { get; set; } = string.Empty;
    }
} 