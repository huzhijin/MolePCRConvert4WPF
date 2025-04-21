using System.Collections.Generic;
using MolePCRConvert4WPF.Core.Models.PanelRules;

namespace MolePCRConvert4WPF.Core.Interfaces
{
    /// <summary>
    /// Panel规则仓库接口
    /// </summary>
    public interface IPanelRuleRepository
    {
        /// <summary>
        /// 获取所有可用的Panel类型名称
        /// </summary>
        /// <returns>Panel类型名称列表</returns>
        List<string> GetAvailablePanelTypeNames();
        
        /// <summary>
        /// 获取指定Panel名称的规则配置
        /// </summary>
        /// <param name="panelName">Panel名称</param>
        /// <returns>Panel规则配置，如果找不到则返回null</returns>
        PanelRuleConfiguration? GetPanelRuleConfiguration(string panelName);
        
        /// <summary>
        /// 获取指定Panel名称和通道名称的通道配置
        /// </summary>
        /// <param name="panelName">Panel名称</param>
        /// <param name="channelName">通道名称</param>
        /// <returns>通道配置，如果找不到则返回null</returns>
        ChannelConfiguration? GetChannelConfiguration(string panelName, string channelName);
        
        /// <summary>
        /// 获取所有Panel规则配置
        /// </summary>
        /// <returns>所有Panel规则配置的列表</returns>
        List<PanelRuleConfiguration> GetAllPanelRuleConfigurations();
        
        /// <summary>
        /// 保存Panel规则配置 (通常保存到文件)
        /// </summary>
        /// <returns>保存是否成功</returns>
        bool SaveChanges();
    }
} 