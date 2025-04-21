namespace MolePCRConvert4WPF.Core.Models
{
    /// <summary>
    /// PCR分析规则模型
    /// </summary>
    public class AnalysisRule
    {
        /// <summary>
        /// 孔位
        /// </summary>
        public string WellPosition { get; set; } = "";
        
        /// <summary>
        /// 荧光通道
        /// </summary>
        public string Channel { get; set; } = "";
        
        /// <summary>
        /// 靶标名称
        /// </summary>
        public string TargetName { get; set; } = "";
        
        /// <summary>
        /// 阳性判定公式
        /// </summary>
        public string PositiveRule { get; set; } = "";
        
        /// <summary>
        /// 浓度计算公式
        /// </summary>
        public string ConcentrationFormula { get; set; } = "";
    }
} 