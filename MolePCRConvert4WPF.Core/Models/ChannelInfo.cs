using System;

namespace MolePCRConvert4WPF.Core.Models
{
    /// <summary>
    /// 通道信息
    /// </summary>
    public class ChannelInfo
    {
        /// <summary>
        /// 通道ID
        /// </summary>
        public Guid Id { get; set; } = Guid.NewGuid();
        
        /// <summary>
        /// 通道名称（例如FAM, VIC, ROX, CY5等）
        /// </summary>
        public string Name { get; set; } = string.Empty;
        
        /// <summary>
        /// 显示名称
        /// </summary>
        public string DisplayName { get; set; } = string.Empty;
        
        /// <summary>
        /// 颜色代码
        /// </summary>
        public string ColorCode { get; set; } = string.Empty;
        
        /// <summary>
        /// 波长（nm）
        /// </summary>
        public int? Wavelength { get; set; }
        
        /// <summary>
        /// 描述
        /// </summary>
        public string Description { get; set; } = string.Empty;
        
        /// <summary>
        /// 是否启用
        /// </summary>
        public bool IsEnabled { get; set; } = true;
    }
} 