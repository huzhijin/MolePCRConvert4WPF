using System;

namespace MolePCRConvert4WPF.Core.Models
{
    /// <summary>
    /// 特殊分析规则模型，用于处理需要多个通道组合分析的情况
    /// </summary>
    public class SpecialRule
    {
        /// <summary>
        /// 孔位名称（如F12）
        /// </summary>
        public string WellPosition { get; set; } = "";
        
        /// <summary>
        /// 组合判断规则（如ABS({FAM}-{VIC})<10）
        /// </summary>
        public string CombinedRule { get; set; } = "";
        
        /// <summary>
        /// 受影响的通道列表
        /// </summary>
        public string[] AffectedChannels { get; set; } = Array.Empty<string>();
    }
} 