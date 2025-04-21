using System;

namespace MolePCRConvert4WPF.Core.Models
{
    /// <summary>
    /// Represents a single rule/row in the Analysis Method configuration.
    /// </summary>
    public class AnalysisMethodRule
    {
        // Properties corresponding to DataGrid columns in the old view
        public int Index { get; set; } // Or maybe Guid Id if needed
        public string? WellPosition { get; set; }
        public string? Channel { get; set; }
        public string? SpeciesName { get; set; }
        
        /// <summary>
        /// 阳性判定公式，与旧版的PositiveExpression兼容
        /// </summary>
        public string? JudgeFormula { get; set; }
        
        /// <summary>
        /// 浓度计算公式，与旧版的ConcentrationExpression兼容
        /// </summary>
        public string? ConcentrationFormula { get; set; }

        // Default constructor
        public AnalysisMethodRule() { }

        /// <summary>
        /// 孔位的匹配模式，支持通配符和范围
        /// </summary>
        public string? WellPositionPattern { get; set; } // e.g., "A1", "B:1-6", "*:12", "*"
        
        /// <summary>
        /// 靶标名称，与旧版的SpeciesName兼容
        /// </summary>
        public string? TargetName { get; set; }
        
        /// <summary>
        /// 阳性判定公式，与JudgeFormula兼容，优先使用
        /// </summary>
        public string? PositiveCutoffFormula 
        { 
            get => !string.IsNullOrWhiteSpace(JudgeFormula) ? JudgeFormula : _positiveCutoffFormula;
            set => _positiveCutoffFormula = value;
        }
        private string? _positiveCutoffFormula;
        
        /// <summary>
        /// 孔位，与WellPosition兼容，用于旧版规则匹配
        /// </summary>
        public string? Hole
        {
            get => WellPosition;
            set => WellPosition = value;
        }

        /// <summary>
        /// Checks if a given well position matches the pattern defined in this rule.
        /// </summary>
        /// <param name="wellPosition">The well position string (e.g., "A1").</param>
        /// <returns>True if matches, False otherwise.</returns>
        public bool WellPositionPatternMatches(string? wellPosition)
        { 
            if (string.IsNullOrWhiteSpace(WellPositionPattern) || string.IsNullOrWhiteSpace(wellPosition))
                return false;
            
            // Handle universal wildcard first
            if (WellPositionPattern == "*") return true;
            
            // Handle specific pattern matching (add more complex logic as needed)
            // Example: Exact match
            if (WellPositionPattern.Equals(wellPosition, StringComparison.OrdinalIgnoreCase)) return true;

            // Example: Row wildcard "A:*"
            if (WellPositionPattern.EndsWith(":*"))
            {
                string patternRow = WellPositionPattern.Split(':')[0];
                string positionRow = wellPosition.Length > 0 ? wellPosition.Substring(0, 1) : string.Empty;
                if (patternRow.Equals(positionRow, StringComparison.OrdinalIgnoreCase)) return true;
            }
            
            // Example: Column wildcard "*:12"
            if (WellPositionPattern.StartsWith("*:"))
            {
                 string patternCol = WellPositionPattern.Split(':')[1];
                 string positionCol = wellPosition.Length > 1 ? wellPosition.Substring(1) : string.Empty;
                 if (patternCol.Equals(positionCol, StringComparison.OrdinalIgnoreCase)) return true;
            }

            // Example: Range matching "B:1-6"
            if (WellPositionPattern.Contains(":") && WellPositionPattern.Contains("-"))
            {
                try 
                {
                    string[] parts = WellPositionPattern.Split(':');
                    string patternRow = parts[0];
                    string positionRow = wellPosition.Length > 0 ? wellPosition.Substring(0, 1) : string.Empty;

                    if(patternRow.Equals(positionRow, StringComparison.OrdinalIgnoreCase))
                    {
                        string[] range = parts[1].Split('-');
                        int startCol = int.Parse(range[0]);
                        int endCol = int.Parse(range[1]);
                        int positionCol = int.Parse(wellPosition.Substring(1));
                        if (positionCol >= startCol && positionCol <= endCol) return true;
                    }
                }
                 catch (Exception ex) { /* Log error or ignore invalid pattern */ }
            }
            
            return false; // No match found
        }
    }
} 