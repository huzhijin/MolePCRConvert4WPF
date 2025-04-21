using System.Collections.Generic;

namespace MolePCRConvert4WPF.Core.Models
{
    /// <summary>
    /// Represents the configuration for a specific analysis method.
    /// </summary>
    public class AnalysisMethodConfiguration
    {
        public List<AnalysisMethodRule> Rules { get; set; } = new List<AnalysisMethodRule>();
        // Add other configuration properties if needed, e.g., Name, Description
        public string? Name { get; set; }
        public string? Description { get; set; }
    }
} 