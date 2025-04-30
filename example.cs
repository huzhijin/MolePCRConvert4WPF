using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using MolePCRConvert4WPF.App.Models;
using MolePCRConvert4WPF.Core.Models;

namespace MolePCRConvert4WPF.App.Utils
{
    public static class TemplateVariableConverter
    {
        /// <summary>
        /// 将按类别分组的变量字典转换为VariableCategoryGroup的ObservableCollection
        /// </summary>
        /// <param name="variablesByCategory">按类别分组的变量字典</param>
        /// <returns>VariableCategoryGroup的ObservableCollection</returns>
        public static ObservableCollection<VariableCategoryGroup> ConvertDictionaryToVariableCategoryGroups(
            Dictionary<string, List<TemplateVariable>> variablesByCategory)
        {
            // 创建新的ObservableCollection用于存储结果
            var result = new ObservableCollection<VariableCategoryGroup>();
            
            // 遍历字典中的每个键值对
            foreach (var category in variablesByCategory)
            {
                // 创建一个新的VariableCategoryGroup
                var categoryGroup = new VariableCategoryGroup
                {
                    Name = category.Key
                };
                
                // 将List<TemplateVariable>转换为ObservableCollection<TemplateVariable>
                foreach (var variable in category.Value)
                {
                    categoryGroup.Variables.Add(variable);
                }
                
                // 将新创建的分组添加到结果集合中
                result.Add(categoryGroup);
            }
            
            return result;
        }
        
        /// <summary>
        /// 使用LINQ简化版本的转换方法
        /// </summary>
        /// <param name="variablesByCategory">按类别分组的变量字典</param>
        /// <returns>VariableCategoryGroup的ObservableCollection</returns>
        public static ObservableCollection<VariableCategoryGroup> ConvertDictionaryToVariableCategoryGroupsLinq(
            Dictionary<string, List<TemplateVariable>> variablesByCategory)
        {
            return new ObservableCollection<VariableCategoryGroup>(
                variablesByCategory.Select(category => new VariableCategoryGroup
                {
                    Name = category.Key,
                    Variables = new ObservableCollection<TemplateVariable>(category.Value)
                })
            );
        }
        
        /// <summary>
        /// 示例用法
        /// </summary>
        public static void Example()
        {
            // 假设我们有一个已分组的模板变量字典
            var variablesByCategory = new Dictionary<string, List<TemplateVariable>>
            {
                {
                    "系统", new List<TemplateVariable>
                    {
                        new TemplateVariable { Name = "${Date}", DisplayName = "当前日期", Category = "系统", Description = "当前日期 (yyyy-MM-dd)" },
                        new TemplateVariable { Name = "${Time}", DisplayName = "当前时间", Category = "系统", Description = "当前时间 (HH:mm:ss)" }
                    }
                },
                {
                    "样本信息", new List<TemplateVariable>
                    {
                        new TemplateVariable { Name = "${SampleCount}", DisplayName = "样本数量", Category = "样本信息", Description = "检测样本总数" },
                        new TemplateVariable { Name = "${PositiveSampleCount}", DisplayName = "阳性样本数", Category = "样本信息", Description = "阳性样本数量" }
                    }
                }
            };
            
            // 转换为ObservableCollection<VariableCategoryGroup>
            var categoryGroups = ConvertDictionaryToVariableCategoryGroups(variablesByCategory);
            
            // 或者使用LINQ简化版本
            var categoryGroupsLinq = ConvertDictionaryToVariableCategoryGroupsLinq(variablesByCategory);
            
            // 输出结果以验证
            Console.WriteLine("转换结果:");
            foreach (var group in categoryGroups)
            {
                Console.WriteLine($"类别: {group.Name}");
                foreach (var variable in group.Variables)
                {
                    Console.WriteLine($"  - {variable.DisplayName}: {variable.Description}");
                }
            }
        }
    }
} 