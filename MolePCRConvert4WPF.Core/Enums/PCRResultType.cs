namespace MolePCRConvert4WPF.Core.Enums
{
    /// <summary>
    /// PCR结果类型
    /// </summary>
    public enum PCRResultType
    {
        /// <summary>
        /// 阳性
        /// </summary>
        Positive,
        
        /// <summary>
        /// 阴性
        /// </summary>
        Negative,
        
        /// <summary>
        /// 无效（未确定）
        /// </summary>
        Invalid,
        
        /// <summary>
        /// 弱阳性
        /// </summary>
        WeakPositive,
        
        /// <summary>
        /// 未检测到（CT值大于截止值）
        /// </summary>
        NotDetected
    }
} 