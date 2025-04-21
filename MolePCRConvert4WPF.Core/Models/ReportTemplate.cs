using System;
using System.ComponentModel.DataAnnotations;

namespace MolePCRConvert4WPF.Core.Models
{
    /// <summary>
    /// 报告模板
    /// </summary>
    public class ReportTemplate
    {
        /// <summary>
        /// 模板ID
        /// </summary>
        [Key]
        public Guid Id { get; set; } = Guid.NewGuid();
        
        /// <summary>
        /// 模板名称
        /// </summary>
        [Required]
        [StringLength(100)]
        public string Name { get; set; } = string.Empty;
        
        /// <summary>
        /// 模板描述
        /// </summary>
        [StringLength(500)]
        public string? Description { get; set; }
        
        /// <summary>
        /// 模板文件路径
        /// </summary>
        [Required]
        [StringLength(500)]
        public string FilePath { get; set; } = string.Empty;
        
        /// <summary>
        /// 语言代码 (例如: zh-CN, en-US)
        /// </summary>
        [StringLength(10)]
        public string? LanguageCode { get; set; }
        
        /// <summary>
        /// 创建时间
        /// </summary>
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
        
        /// <summary>
        /// 创建人ID
        /// </summary>
        public Guid? CreatedById { get; set; }
        
        /// <summary>
        /// 更新时间
        /// </summary>
        public DateTime? UpdatedAt { get; set; }
        
        /// <summary>
        /// 更新人ID
        /// </summary>
        public Guid? UpdatedById { get; set; }
        
        /// <summary>
        /// 是否为Word模板
        /// </summary>
        public bool IsWordTemplate { get; set; } = false;
        
        /// <summary>
        /// 是否为Excel模板
        /// </summary>
        public bool IsExcelTemplate { get; set; } = true;
        
        /// <summary>
        /// 是否为HTML模板
        /// </summary>
        public bool IsHtmlTemplate { get; set; } = false;
        
        /// <summary>
        /// 是否为系统预设模板
        /// </summary>
        public bool IsSystem { get; set; } = false;
        
        /// <summary>
        /// 是否启用
        /// </summary>
        public bool IsEnabled { get; set; } = true;
    }
} 