using System.Collections.Generic;
using System.Collections.ObjectModel;
using MolePCRConvert4WPF.Core.Models;

namespace MolePCRConvert4WPF.App.Models
{
    /// <summary>
    /// 变量分类组，用于在UI中按分类显示变量
    /// </summary>
    public class VariableCategoryGroup
    {
        /// <summary>
        /// 分类名称
        /// </summary>
        public string CategoryName { get; set; }

        /// <summary>
        /// 该分类下的变量集合
        /// </summary>
        public ObservableCollection<MolePCRConvert4WPF.Core.Models.TemplateVariable> Variables { get; set; } = new ObservableCollection<MolePCRConvert4WPF.Core.Models.TemplateVariable>();
    }
} 