using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using MolePCRConvert4WPF.Core.Enums;

namespace MolePCRConvert4WPF.Core.Models
{
    /// <summary>
    /// 板信息
    /// </summary>
    public class Plate
    {
        /// <summary>
        /// 板ID
        /// </summary>
        [Key]
        public Guid Id { get; set; } = Guid.NewGuid();
        
        /// <summary>
        /// 板名称
        /// </summary>
        [Required]
        [StringLength(100)]
        public string Name { get; set; } = string.Empty;
        
        /// <summary>
        /// 仪器类型
        /// </summary>
        public InstrumentType InstrumentType { get; set; }
        
        /// <summary>
        /// 创建时间
        /// </summary>
        public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
        
        /// <summary>
        /// 创建人ID
        /// </summary>
        public Guid? CreatedById { get; set; }
        
        /// <summary>
        /// 行数
        /// </summary>
        public int Rows { get; set; } = 8; // 默认8行 (A-H)
        
        /// <summary>
        /// 列数
        /// </summary>
        public int Columns { get; set; } = 12; // 默认12列
        
        /// <summary>
        /// 运行ID/批次号
        /// </summary>
        [StringLength(50)]
        public string? RunId { get; set; }
        
        /// <summary>
        /// 分析方法ID
        /// </summary>
        public Guid? AnalysisMethodId { get; set; }
        
        /// <summary>
        /// 分析方法名称
        /// </summary>
        [StringLength(100)]
        public string? AnalysisMethod { get; set; }
        
        /// <summary>
        /// 导入文件路径
        /// </summary>
        [StringLength(500)]
        public string? ImportFilePath { get; set; }
        
        /// <summary>
        /// 备注
        /// </summary>
        [StringLength(500)]
        public string? Notes { get; set; }
        
        /// <summary>
        /// 是否已分析
        /// </summary>
        public bool IsAnalyzed { get; set; } = false;
        
        /// <summary>
        /// 分析时间
        /// </summary>
        public DateTime? AnalyzedAt { get; set; }
        
        /// <summary>
        /// 分析人ID
        /// </summary>
        public Guid? AnalyzedById { get; set; }
        
        /// <summary>
        /// 孔位布局集合
        /// </summary>
        public List<WellLayout> WellLayouts { get; set; } = new List<WellLayout>();
        
        /// <summary>
        /// 样本集合
        /// </summary>
        public List<Sample> Samples { get; set; } = new List<Sample>();
        
        /// <summary>
        /// 通道信息集合
        /// </summary>
        public List<ChannelInfo> ChannelInfos { get; set; } = new List<ChannelInfo>();
    }
} 