using System;
using System.Collections.Generic;
using System.Runtime.Serialization;

namespace MolePCRConvert4WPF.Core.Models
{
    /// <summary>
    /// 自定义报告模板
    /// </summary>
    [DataContract]
    public class ReportCustomTemplate
    {
        /// <summary>
        /// 模板名称
        /// </summary>
        [DataMember]
        public string Name { get; set; }

        /// <summary>
        /// 模板创建时间
        /// </summary>
        [DataMember]
        public DateTime CreatedTime { get; set; }

        /// <summary>
        /// 最后修改时间
        /// </summary>
        [DataMember]
        public DateTime LastModifiedTime { get; set; }

        /// <summary>
        /// 模板描述
        /// </summary>
        [DataMember]
        public string Description { get; set; }

        /// <summary>
        /// 模板内容（单元格数据）
        /// </summary>
        [DataMember]
        public List<TemplateCellData> Cells { get; set; } = new List<TemplateCellData>();

        /// <summary>
        /// 数据表格标记位置
        /// </summary>
        [DataMember]
        public TemplatePosition DataTablePosition { get; set; }

        /// <summary>
        /// 列数
        /// </summary>
        [DataMember]
        public int ColumnCount { get; set; }

        /// <summary>
        /// 行数
        /// </summary>
        [DataMember]
        public int RowCount { get; set; }
    }

    /// <summary>
    /// 模板单元格数据
    /// </summary>
    [DataContract]
    public class TemplateCellData
    {
        /// <summary>
        /// 行索引
        /// </summary>
        [DataMember]
        public int RowIndex { get; set; }

        /// <summary>
        /// 列索引
        /// </summary>
        [DataMember]
        public int ColumnIndex { get; set; }

        /// <summary>
        /// 单元格内容
        /// </summary>
        [DataMember]
        public string Content { get; set; }

        /// <summary>
        /// 背景颜色
        /// </summary>
        [DataMember]
        public string BackgroundColor { get; set; }

        /// <summary>
        /// 前景颜色
        /// </summary>
        [DataMember]
        public string ForegroundColor { get; set; }

        /// <summary>
        /// 字体名称
        /// </summary>
        [DataMember]
        public string FontName { get; set; }

        /// <summary>
        /// 字体大小
        /// </summary>
        [DataMember]
        public float FontSize { get; set; }

        /// <summary>
        /// 是否粗体
        /// </summary>
        [DataMember]
        public bool IsBold { get; set; }

        /// <summary>
        /// 是否斜体
        /// </summary>
        [DataMember]
        public bool IsItalic { get; set; }

        /// <summary>
        /// 单元格合并列数
        /// </summary>
        [DataMember]
        public int ColumnSpan { get; set; } = 1;

        /// <summary>
        /// 单元格合并行数
        /// </summary>
        [DataMember]
        public int RowSpan { get; set; } = 1;

        /// <summary>
        /// 水平对齐方式
        /// </summary>
        [DataMember]
        public CellHorizontalAlignment HorizontalAlignment { get; set; } = CellHorizontalAlignment.Left;

        /// <summary>
        /// 垂直对齐方式
        /// </summary>
        [DataMember]
        public CellVerticalAlignment VerticalAlignment { get; set; } = CellVerticalAlignment.Center;
    }

    /// <summary>
    /// 模板中的位置
    /// </summary>
    [DataContract]
    public class TemplatePosition
    {
        /// <summary>
        /// 行索引
        /// </summary>
        [DataMember]
        public int RowIndex { get; set; }

        /// <summary>
        /// 列索引
        /// </summary>
        [DataMember]
        public int ColumnIndex { get; set; }
    }

    /// <summary>
    /// 单元格水平对齐方式
    /// </summary>
    public enum CellHorizontalAlignment
    {
        Left,
        Center,
        Right
    }

    /// <summary>
    /// 单元格垂直对齐方式
    /// </summary>
    public enum CellVerticalAlignment
    {
        Top,
        Center,
        Bottom
    }
} 