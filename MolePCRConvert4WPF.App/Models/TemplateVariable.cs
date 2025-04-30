using System;

namespace MolePCRConvert4WPF.App.Models
{
    /// <summary>
    /// 模板变量类，用于报告模板中使用的变量
    /// </summary>
    public class TemplateVariable
    {
        /// <summary>
        /// 变量名称，用于在模板中使用如 ${VariableName}
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 显示名称，用于在UI界面上显示
        /// </summary>
        public string DisplayName { get; set; }

        /// <summary>
        /// 变量描述，用于提示用户变量的用途
        /// </summary>
        public string Description { get; set; }

        /// <summary>
        /// 变量所属分类
        /// </summary>
        public string Category { get; set; }
    }
} 