using System;
using System.Collections.Generic;
using MolePCRConvert4WPF.Core.Models;
using MolePCRConvert4WPF.Core.Models.PanelRules;

namespace MolePCRConvert4WPF.Core.Interfaces
{
    /// <summary>
    /// 样本命名服务接口
    /// </summary>
    public interface ISampleNamingService
    {
        /// <summary>
        /// 根据Panel类型生成分析方法感知的样本命名
        /// </summary>
        /// <param name="panelType">Panel类型</param>
        /// <param name="rows">孔板行数，默认8</param>
        /// <param name="columns">孔板列数，默认12</param>
        /// <returns>样本-孔位映射列表</returns>
        List<SampleWellMapping> GenerateAnalysisAwareNames(string panelType, int rows = 8, int columns = 12);
        
        /// <summary>
        /// 获取指定Panel类型对应的样本单元孔数
        /// </summary>
        /// <param name="panelType">Panel类型</param>
        /// <returns>每个样本的孔数</returns>
        int GetSampleUnitWellCount(string panelType);
        
        /// <summary>
        /// 手动命名孔位时验证是否符合分析方法规则
        /// </summary>
        /// <param name="panelType">Panel类型</param>
        /// <param name="wellPosition">孔位位置</param>
        /// <param name="sampleName">样本名称</param>
        /// <returns>验证结果和错误信息</returns>
        (bool IsValid, string ErrorMessage) ValidateWellNaming(string panelType, string wellPosition, string sampleName);
        
        /// <summary>
        /// 创建孔位到样本和反应体系的映射表
        /// </summary>
        /// <param name="plateId">板ID</param>
        /// <param name="panelType">Panel类型</param>
        /// <param name="wells">孔位布局列表</param>
        /// <returns>样本-孔位映射列表</returns>
        List<SampleWellMapping> CreateSampleWellMappings(Guid plateId, string panelType, List<WellLayout> wells);
        
        /// <summary>
        /// 标记内标样本
        /// </summary>
        /// <param name="mappings">样本-孔位映射列表</param>
        /// <param name="panelType">Panel类型</param>
        /// <returns>更新后的映射列表</returns>
        List<SampleWellMapping> MarkInternalControlSamples(List<SampleWellMapping> mappings, string panelType);
        
        /// <summary>
        /// 根据特定前缀标记样本类型（如"靶标_"、"阳性对照_"等）
        /// </summary>
        /// <param name="mappings">样本-孔位映射列表</param>
        /// <returns>更新后的映射列表</returns>
        List<SampleWellMapping> MarkSamplesByPrefix(List<SampleWellMapping> mappings);
    }
} 