using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel.DataAnnotations;
using MolePCRConvert4WPF.Core.Models.PanelRules; // Assuming TargetRuleModel is defined here or needs its own file

namespace MolePCRConvert4WPF.Core.Models
{
    /// <summary>
    /// 分析方法
    /// </summary>
    public class AnalysisMethod
    {
        /// <summary>
        /// 分析方法ID
        /// </summary>
        [Key]
        public Guid Id { get; set; } = Guid.NewGuid();
        
        /// <summary>
        /// 方法名称
        /// </summary>
        [Required]
        [StringLength(100)]
        public string Name { get; set; } = string.Empty;
        
        /// <summary>
        /// 方法描述
        /// </summary>
        [StringLength(500)]
        public string? Description { get; set; }
        
        /// <summary>
        /// 规则JSON (可能不再需要，如果使用下面的模型)
        /// </summary>
        // public string RulesJson { get; set; } = string.Empty;
        
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
        /// 是否为系统预设方法
        /// </summary>
        public bool IsSystem { get; set; } = false;
        
        /// <summary>
        /// 是否启用
        /// </summary>
        public bool IsEnabled { get; set; } = true;
        
        /// <summary>
        /// 微生物孔分析方法
        /// </summary>
        public MicrobeHoleAnalyticalMethodModel? MicrobeHoleAnalyticalMethod { get; set; }
        
        /// <summary>
        /// 分析方法Excel模板路径
        /// </summary>
        [StringLength(500)]
        public string? TemplateFilePath { get; set; }
    }
    
    /// <summary>
    /// 微生物孔分析方法模型 (用于Excel配置)
    /// </summary>
    public class MicrobeHoleAnalyticalMethodModel
    {
        /// <summary>
        /// 方法名称
        /// </summary>
        public string Name { get; set; } = string.Empty;
        
        /// <summary>
        /// 方法描述
        /// </summary>
        public string Description { get; set; } = string.Empty;
        
        /// <summary>
        /// 版本
        /// </summary>
        public string Version { get; set; } = "1.0";
        
        /// <summary>
        /// 创建时间
        /// </summary>
        public DateTime CreatedTime { get; set; } = DateTime.Now;
        
        /// <summary>
        /// 更新时间
        /// </summary>
        public DateTime ModifiedTime { get; set; } = DateTime.Now;
        
        /// <summary>
        /// 更新者
        /// </summary>
        public string ModifiedBy { get; set; } = string.Empty;
        
        /// <summary>
        /// 配置属性
        /// </summary>
        public AnalytConfigProperties AnalytConfigProperties { get; set; } = new AnalytConfigProperties();
        
        /// <summary>
        /// 孔位-分析表达式映射列表
        /// </summary>
        public List<WellAnalytExpressionMapping> WellAnalytExpressionMappingList { get; set; } = new List<WellAnalytExpressionMapping>();
        
        /// <summary>
        /// 靶标规则列表 (可能与 PanelRules/Rule 重复，需确认)
        /// </summary>
        public ObservableCollection<TargetRuleModel> TargetRules { get; set; } = new ObservableCollection<TargetRuleModel>();
    }
    
    /// <summary>
    /// 分析配置属性
    /// </summary>
    public class AnalytConfigProperties
    {
        /// <summary>
        /// 默认阳性表达式
        /// </summary>
        public string DefaultPositiveExpression { get; set; } = "{CT}<35";
        
        /// <summary>
        /// 支持人数
        /// </summary>
        public int SupportedPersonCount { get; set; } = 12;
        
        /// <summary>
        /// 其他配置项
        /// </summary>
        public Dictionary<string, string> OtherProperties { get; set; } = new Dictionary<string, string>();
    }
    
    /// <summary>
    /// 孔位-分析表达式映射
    /// </summary>
    public class WellAnalytExpressionMapping
    {
        /// <summary>
        /// 孔位（如A1, B2等）
        /// </summary>
        public string Hole { get; set; } = string.Empty;
        
        /// <summary>
        /// 通道（如FAM, HEX等）
        /// </summary>
        public string Channel { get; set; } = string.Empty;
        
        /// <summary>
        /// 微生物/靶基因名称
        /// </summary>
        public string MicrobeOrGenName { get; set; } = string.Empty;
        
        /// <summary>
        /// 阳性判定表达式（如{CT}<35）
        /// </summary>
        public string PositiveExpression { get; set; } = string.Empty;
        
        /// <summary>
        /// 浓度计算表达式（如10^{20-CT}）
        /// </summary>
        public string ConcentrationExpression { get; set; } = string.Empty;
        
        /// <summary>
        /// 自定义属性
        /// </summary>
        public Dictionary<string, object> CustomProperties { get; set; } = new Dictionary<string, object>();
    }

    /// <summary>
    /// 靶标规则模型 (需要定义或确认其来源)
    /// </summary>
    public class TargetRuleModel
    {
        public Guid Id { get; set; } = Guid.NewGuid();
        public string Name { get; set; } = string.Empty;
        public string Condition { get; set; } = string.Empty;
        public string Result { get; set; } = string.Empty;
        // ... other properties if needed
    }
} 