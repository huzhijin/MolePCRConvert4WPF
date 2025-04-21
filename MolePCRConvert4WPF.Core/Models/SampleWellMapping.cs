using System;
using System.Collections.Generic;

namespace MolePCRConvert4WPF.Core.Models
{
    /// <summary>
    /// 样本-孔位映射
    /// </summary>
    public class SampleWellMapping
    {
        /// <summary>
        /// 样本名称
        /// </summary>
        public string? SampleName { get; set; }
        
        /// <summary>
        /// 样本ID
        /// </summary>
        public Guid SampleId { get; set; }
        
        /// <summary>
        /// 板ID
        /// </summary>
        public Guid PlateId { get; set; }
        
        /// <summary>
        /// 关联的孔位位置列表
        /// </summary>
        public List<string> WellPositions { get; set; } = new List<string>();
        
        /// <summary>
        /// 关联的孔位ID列表
        /// </summary>
        public List<Guid> WellIds { get; set; } = new List<Guid>();
        
        /// <summary>
        /// 是否为内标
        /// </summary>
        public bool IsInternalControl { get; set; } = false;
        
        /// <summary>
        /// 是否为阳性对照
        /// </summary>
        public bool IsPositiveControl { get; set; } = false;
        
        /// <summary>
        /// 是否为阴性对照
        /// </summary>
        public bool IsNegativeControl { get; set; } = false;
        
        /// <summary>
        /// 是否为标准品
        /// </summary>
        public bool IsStandard { get; set; } = false;
        
        /// <summary>
        /// 分析结果
        /// </summary>
        public string? AnalysisResult { get; set; }
        
        /// <summary>
        /// CT值
        /// </summary>
        public Dictionary<string, double?> CtValues { get; set; } = new Dictionary<string, double?>();
        
        /// <summary>
        /// 患者信息
        /// </summary>
        public PatientInfo? PatientInfo { get; set; }
    }
} 