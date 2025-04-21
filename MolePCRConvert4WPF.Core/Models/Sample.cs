using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using MolePCRConvert4WPF.Core.Enums;

namespace MolePCRConvert4WPF.Core.Models
{
    /// <summary>
    /// 样本信息
    /// </summary>
    public class Sample
    {
        /// <summary>
        /// 样本ID
        /// </summary>
        [Key]
        public Guid Id { get; set; } = Guid.NewGuid();
        
        /// <summary>
        /// 样本名称
        /// </summary>
        [Required]
        [StringLength(100)]
        public string Name { get; set; } = string.Empty;
        
        /// <summary>
        /// 板ID
        /// </summary>
        [Required]
        public Guid PlateId { get; set; }
        
        /// <summary>
        /// 导入文件ID
        /// </summary>
        public Guid? ImportFileId { get; set; }
        
        /// <summary>
        /// 样本编号
        /// </summary>
        [StringLength(50)]
        public string? SampleCode { get; set; }
        
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
        /// 样本结果
        /// </summary>
        public SampleResult Result { get; set; } = SampleResult.Unknown;
        
        /// <summary>
        /// 浓度值
        /// </summary>
        public double? Concentration { get; set; }
        
        /// <summary>
        /// 浓度单位
        /// </summary>
        [StringLength(20)]
        public string? ConcentrationUnit { get; set; }
        
        /// <summary>
        /// 备注
        /// </summary>
        [StringLength(500)]
        public string? Notes { get; set; }
        
        /// <summary>
        /// 相关孔位
        /// </summary>
        [NotMapped]
        public List<WellLayout> Wells { get; set; } = new List<WellLayout>();
    }
} 