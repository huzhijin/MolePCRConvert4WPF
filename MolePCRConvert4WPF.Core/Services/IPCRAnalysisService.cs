using MolePCRConvert4WPF.Core.Models;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace MolePCRConvert4WPF.Core.Services
{
    /// <summary>
    /// Defines the contract for the PCR analysis service.
    /// </summary>
    public interface IPCRAnalysisService
    {
        /// <summary>
        /// Analyzes the provided plate data using the specified configuration.
        /// </summary>
        /// <param name="plateData">The plate data containing well layouts, patient info, and potentially raw Ct values.</param>
        /// <param name="analysisConfig">The analysis method configuration containing rules and formulas.</param>
        /// <returns>A list of analysis result items.</returns>
        Task<List<AnalysisResultItem>> AnalyzeAsync(Plate plateData, AnalysisMethodConfiguration analysisConfig);
        // Adjust AnalysisMethodConfiguration type if it differs
    }
} 