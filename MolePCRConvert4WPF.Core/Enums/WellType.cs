using System;

namespace MolePCRConvert4WPF.Core.Enums
{
    /// <summary>
    /// 孔位类型
    /// </summary>
    public enum WellType
    {
        /// <summary>
        /// 样本
        /// </summary>
        Sample,
        
        /// <summary>
        /// 阳性对照
        /// </summary>
        PositiveControl,
        
        /// <summary>
        /// 阴性对照
        /// </summary>
        NegativeControl,
        
        /// <summary>
        /// 标准品
        /// </summary>
        Standard,
        
        /// <summary>
        /// 空白
        /// </summary>
        Empty,
        
        /// <summary>
        /// 未知类型
        /// </summary>
        Unknown,
        
        /// <summary>
        /// 内标
        /// </summary>
        InternalControl,
        
        /// <summary>
        /// 未使用
        /// </summary>
        Unused
    }
} 