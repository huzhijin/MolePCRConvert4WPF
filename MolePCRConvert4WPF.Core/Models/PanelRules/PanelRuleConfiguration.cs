using System;
using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace MolePCRConvert4WPF.Core.Models.PanelRules
{
    /// <summary>
    /// 面板规则配置
    /// </summary>
    public class PanelRuleConfiguration
    {
        /// <summary>
        /// 规则ID
        /// </summary>
        public Guid Id { get; set; } = Guid.NewGuid();
        
        /// <summary>
        /// 规则名称
        /// </summary>
        public string Name { get; set; } = string.Empty;
        
        /// <summary>
        /// 规则描述
        /// </summary>
        public string Description { get; set; } = string.Empty;
        
        /// <summary>
        /// 版本号
        /// </summary>
        public string Version { get; set; } = "1.0";
        
        /// <summary>
        /// 最后更新时间
        /// </summary>
        public DateTime LastUpdated { get; set; } = DateTime.UtcNow;
        
        /// <summary>
        /// 面板规则列表
        /// </summary>
        public List<RuleGroup> RuleGroups { get; set; } = new List<RuleGroup>();
        
        /// <summary>
        /// 通道配置
        /// </summary>
        public List<ChannelConfiguration> Channels { get; set; } = new List<ChannelConfiguration>();
    }
    
    /// <summary>
    /// 规则组
    /// </summary>
    public class RuleGroup
    {
        /// <summary>
        /// 规则组ID
        /// </summary>
        public Guid Id { get; set; } = Guid.NewGuid();
        
        /// <summary>
        /// 规则组名称
        /// </summary>
        public string Name { get; set; } = string.Empty;
        
        /// <summary>
        /// 匹配优先级
        /// </summary>
        public int Priority { get; set; } = 0;
        
        /// <summary>
        /// 规则列表
        /// </summary>
        public List<Rule> Rules { get; set; } = new List<Rule>();
    }
    
    /// <summary>
    /// 规则
    /// </summary>
    public class Rule
    {
        /// <summary>
        /// 规则ID
        /// </summary>
        public Guid Id { get; set; } = Guid.NewGuid();
        
        /// <summary>
        /// 规则名称
        /// </summary>
        public string Name { get; set; } = string.Empty;
        
        /// <summary>
        /// 条件表达式
        /// </summary>
        public string Condition { get; set; } = string.Empty;
        
        /// <summary>
        /// 动作
        /// </summary>
        public string Action { get; set; } = string.Empty;
    }
    
    /// <summary>
    /// 通道配置
    /// </summary>
    public class ChannelConfiguration
    {
        /// <summary>
        /// 通道名称
        /// </summary>
        public string Name { get; set; } = string.Empty;
        
        /// <summary>
        /// 目标基因/探针
        /// </summary>
        public string Target { get; set; } = string.Empty;
        
        /// <summary>
        /// 阳性判定最小CT值
        /// </summary>
        public double MinPositiveCt { get; set; } = 10.0;
        
        /// <summary>
        /// 阳性判定最大CT值
        /// </summary>
        public double MaxPositiveCt { get; set; } = 35.0;
    }
} 