using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using MolePCRConvert4WPF.App.Models;
using MolePCRConvert4WPF.Core.Models;

namespace MolePCRConvert4WPF.App.Utils
{
    /// <summary>
    /// 模板变量转换器，用于将不同格式的模板变量集合相互转换
    /// </summary>
    public static class TemplateVariableConverter
    {
        /// <summary>
        /// 将按类别分组的变量字典转换为ObservableCollection&lt;VariableCategoryGroup&gt;
        /// </summary>
        /// <param name="variablesByCategory">按类别分组的变量字典</param>
        /// <returns>可观察的变量类别组集合</returns>
        public static ObservableCollection<VariableCategoryGroup> ConvertDictionaryToVariableCategoryGroups(
            Dictionary<string, List<MolePCRConvert4WPF.Core.Models.TemplateVariable>> variablesByCategory)
        {
            var result = new ObservableCollection<VariableCategoryGroup>();
            
            foreach (var category in variablesByCategory.Keys)
            {
                var group = new VariableCategoryGroup
                {
                    CategoryName = category,
                    Variables = new ObservableCollection<MolePCRConvert4WPF.Core.Models.TemplateVariable>(variablesByCategory[category])
                };
                
                result.Add(group);
            }
            
            return result;
        }
        
        /// <summary>
        /// 使用LINQ将按类别分组的变量字典转换为ObservableCollection&lt;VariableCategoryGroup&gt;
        /// </summary>
        /// <param name="variablesByCategory">按类别分组的变量字典</param>
        /// <returns>可观察的变量类别组集合</returns>
        public static ObservableCollection<VariableCategoryGroup> ConvertDictionaryToVariableCategoryGroupsLinq(
            Dictionary<string, List<MolePCRConvert4WPF.Core.Models.TemplateVariable>> variablesByCategory)
        {
            return new ObservableCollection<VariableCategoryGroup>(
                variablesByCategory.Select(kvp => new VariableCategoryGroup
                {
                    CategoryName = kvp.Key,
                    Variables = new ObservableCollection<MolePCRConvert4WPF.Core.Models.TemplateVariable>(kvp.Value)
                }).OrderBy(g => g.CategoryName)
            );
        }
    }
} 