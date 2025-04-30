using System;

namespace MolePCRConvert4WPF.Core.Models
{
    /// <summary>
    /// 表示报告模板中可用的变量
    /// </summary>
    public class TemplateVariable
    {
        /// <summary>
        /// 变量名称，用于模板中的实际替换
        /// </summary>
        public string Name { get; set; }
        
        /// <summary>
        /// 变量显示名称，用于UI显示
        /// </summary>
        public string DisplayName { get; set; }
        
        /// <summary>
        /// 变量所属类别
        /// </summary>
        public string Category { get; set; }
        
        /// <summary>
        /// 变量的描述信息
        /// </summary>
        public string Description { get; set; }
    }
} 