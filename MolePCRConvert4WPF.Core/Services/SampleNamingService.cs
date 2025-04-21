using System;
using System.Collections.Generic;
using System.Linq;
using MolePCRConvert4WPF.Core.Interfaces;
using MolePCRConvert4WPF.Core.Models;
using MolePCRConvert4WPF.Core.Models.PanelRules;

namespace MolePCRConvert4WPF.Core.Services
{
    /// <summary>
    /// 样本命名服务实现
    /// </summary>
    public class SampleNamingService : ISampleNamingService
    {
        private readonly IPanelRuleRepository _panelRuleRepository;
        
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="panelRuleRepository">Panel规则仓库</param>
        public SampleNamingService(IPanelRuleRepository panelRuleRepository)
        {
            _panelRuleRepository = panelRuleRepository ?? throw new ArgumentNullException(nameof(panelRuleRepository));
        }
        
        // Note: The original implementation relied heavily on panel rules from IPanelRuleRepository
        // which had methods like GetWellsPerSample, GetReactionSystem, IsInternalControlWell, GetInternalControlSystems.
        // The new IPanelRuleRepository focuses on loading/providing PanelRuleConfiguration.
        // We need to adapt the logic here to use the PanelRuleConfiguration object.

        /// <summary>
        /// 根据Panel类型生成分析方法感知的样本命名 (需要重构以适应新的IPanelRuleRepository)
        /// </summary>
        public List<SampleWellMapping> GenerateAnalysisAwareNames(string panelName, int rows = 8, int columns = 12)
        {
            var config = _panelRuleRepository.GetPanelRuleConfiguration(panelName);
            if (config == null)
            {
                // Handle case where panel configuration is not found
                // Perhaps return an empty list or throw an exception
                return new List<SampleWellMapping>(); 
            }
            
            // TODO: Implement logic based on config.RuleGroups, config.Channels, etc.
            // This requires understanding how panel rules define sample grouping, internal controls, etc.
            // For now, returning a basic placeholder mapping.
            
            var mappings = new List<SampleWellMapping>();
            for (int i = 0; i < columns; i++)
            {
                 mappings.Add(new SampleWellMapping
                 {
                      SampleId = Guid.NewGuid(), // Or generate based on column index
                      SampleName = $"样本{i + 1}",
                      PlateId = Guid.Empty, // Needs actual PlateId
                      WellPositions = Enumerable.Range(0, rows).Select(r => $"{(char)('A' + r)}{i + 1}").ToList(),
                      WellIds = Enumerable.Range(0, rows).Select(_ => Guid.NewGuid()).ToList()
                 });
            }
            return mappings;
        }
        
        /// <summary>
        /// 获取指定Panel类型对应的样本单元孔数 (需要重构)
        /// </summary>
        public int GetSampleUnitWellCount(string panelName)
        {
            var config = _panelRuleRepository.GetPanelRuleConfiguration(panelName);
            // TODO: Determine wells per sample based on configuration (this might not be explicit)
            // Placeholder value
            return 1; 
        }
        
        /// <summary>
        /// 手动命名孔位时验证是否符合分析方法规则 (需要重构)
        /// </summary>
        public (bool IsValid, string ErrorMessage) ValidateWellNaming(string panelName, string wellPosition, string sampleName)
        {
            var config = _panelRuleRepository.GetPanelRuleConfiguration(panelName);
            if (config == null) return (false, "未找到对应的Panel规则");
            
            // TODO: Implement validation based on panel rules
            // Placeholder: always valid
            return (true, string.Empty);
        }
        
        /// <summary>
        /// 创建孔位到样本和反应体系的映射表 (需要重构)
        /// </summary>
        public List<SampleWellMapping> CreateSampleWellMappings(Guid plateId, string panelName, List<WellLayout> wells)
        {
            var config = _panelRuleRepository.GetPanelRuleConfiguration(panelName);
            if (config == null) return new List<SampleWellMapping>();
            
            var mappings = new Dictionary<string, SampleWellMapping>();
            
            foreach (var well in wells)
            {
                // Use null-coalescing for sample name
                string sampleKey = well.SampleName ?? $"未知样本_{well.WellName}"; // Use a unique key if name is null
                
                if (!mappings.TryGetValue(sampleKey, out var mapping))
                {
                    mapping = new SampleWellMapping
                    {
                        SampleId = well.SampleId ?? Guid.NewGuid(), // Use existing or generate new
                        SampleName = well.SampleName, // Can be null here
                        PlateId = plateId,
                        // PatientInfo = well.Sample?.PatientInfo // CS1061: Commented out - Sample doesn't have PatientInfo. TODO: Implement correct PatientInfo association logic.
                    };
                    mappings[sampleKey] = mapping;
                }
                mapping.WellPositions.Add(well.WellName);
                mapping.WellIds.Add(well.Id);

                // Add CtValue for the specific channel if available
                if (!string.IsNullOrEmpty(well.Channel) && well.CtValue.HasValue)
                {
                    mapping.CtValues[well.Channel] = well.CtValue;
                }
                
                // Determine control types based on well.WellType
                mapping.IsInternalControl = well.WellType == Enums.WellType.InternalControl;
                mapping.IsPositiveControl = well.WellType == Enums.WellType.PositiveControl;
                mapping.IsNegativeControl = well.WellType == Enums.WellType.NegativeControl;
                mapping.IsStandard = well.WellType == Enums.WellType.Standard;
            }
            
            return MarkSamplesByPrefix(mappings.Values.ToList());
        }
        
        /// <summary>
        /// 标记内标样本 (可能需要调整)
        /// </summary>
        public List<SampleWellMapping> MarkInternalControlSamples(List<SampleWellMapping> mappings, string panelName)
        {
             // This logic might need adjustment based on how Internal Controls are defined in rules.
             // The previous implementation checked reaction systems, the new one checks WellType.
             foreach (var mapping in mappings)
             {
                 if(mapping.IsInternalControl) { // Check the flag set in CreateSampleWellMappings
                    // Optionally add notes or perform other actions
                 }
             }
            return mappings;
        }
        
        /// <summary>
        /// 根据特定前缀标记样本类型 (可能不需要，如果WellType足够)
        /// </summary>
        public List<SampleWellMapping> MarkSamplesByPrefix(List<SampleWellMapping> mappings)
        {
            // Consider if this logic is still needed or if WellType is sufficient.
            // Keeping original logic for now.
            var prefixTypeMap = new Dictionary<string, string>
            {
                { "靶标_", "靶标样本" },
                { "阳性对照_", "阳性对照" },
                { "阴性对照_", "阴性对照" },
                { "标准品_", "标准品" },
                { "内标_", "内标样本" }
            };
            
            foreach (var mapping in mappings)
            {
                foreach (var prefix in prefixTypeMap.Keys)
                {
                    // CS8602 Fix: Add null check for SampleName
                    if (mapping.SampleName != null && mapping.SampleName.StartsWith(prefix))
                    {
                        // Update sample type flags based on prefix if needed
                        if (prefix == "阳性对照_") mapping.IsPositiveControl = true;
                        if (prefix == "阴性对照_") mapping.IsNegativeControl = true;
                        if (prefix == "标准品_") mapping.IsStandard = true;
                        if (prefix == "内标_") mapping.IsInternalControl = true;
                        break;
                    }
                }
            }
            return mappings;
        }
    }
} 