using System;
using unvell.ReoGrid;
using unvell.ReoGrid.Graphics;

namespace MolePCRConvert4WPF.Infrastructure.Extensions
{
    /// <summary>
    /// ReoGrid Worksheet扩展方法
    /// </summary>
    public static class ReoGridExtensions
    {
        /// <summary>
        /// 获取指定范围的样式信息
        /// </summary>
        /// <param name="worksheet">工作表</param>
        /// <param name="range">单元格范围字符串 (例如 "A1:C3")</param>
        /// <returns>范围样式对象</returns>
        public static WorksheetRangeStyle GetRangeStyle(this Worksheet worksheet, string range)
        {
            var rangePos = new RangePosition(range);
            return GetRangeStyle(worksheet, rangePos);
        }

        /// <summary>
        /// 获取指定范围的样式信息
        /// </summary>
        /// <param name="worksheet">工作表</param>
        /// <param name="row">起始行索引</param>
        /// <param name="col">起始列索引</param>
        /// <param name="rows">行数</param>
        /// <param name="cols">列数</param>
        /// <returns>范围样式对象</returns>
        public static WorksheetRangeStyle GetRangeStyle(this Worksheet worksheet, int row, int col, int rows, int cols)
        {
            var rangePos = new RangePosition(row, col, rows, cols);
            return GetRangeStyle(worksheet, rangePos);
        }

        /// <summary>
        /// 获取指定范围的样式信息
        /// </summary>
        /// <param name="worksheet">工作表</param>
        /// <param name="range">单元格范围位置</param>
        /// <returns>范围样式对象</returns>
        public static WorksheetRangeStyle GetRangeStyle(this Worksheet worksheet, RangePosition range)
        {
            var style = new WorksheetRangeStyle();
            
            // 获取范围内第一个单元格的样式作为基准
            var cell = worksheet.GetCell(range.Row, range.Col);
            if (cell != null)
            {
                // 字体信息
                style.FontName = cell.Style.FontName;
                style.FontSize = cell.Style.FontSize;
                
                // 文本格式
                style.IsBold = cell.Style.Bold;
                style.IsItalic = cell.Style.Italic;
                style.IsUnderline = cell.Style.Underline;
                style.TextColor = cell.Style.TextColor;
                
                // 背景色
                style.BackColor = cell.Style.BackColor;
                
                // 对齐方式
                style.HAlign = cell.Style.HAlign;
                style.VAlign = cell.Style.VAlign;
            }
            
            return style;
        }
    }
    
    /// <summary>
    /// 工作表范围样式，用于获取单元格区域的样式信息
    /// </summary>
    public class WorksheetRangeStyle
    {
        public string FontName { get; set; } = "Arial";
        public float FontSize { get; set; } = 10.5f;
        public bool IsBold { get; set; }
        public bool IsItalic { get; set; }
        public bool IsUnderline { get; set; }
        public System.Drawing.Color TextColor { get; set; } = System.Drawing.Color.Black;
        public System.Drawing.Color BackColor { get; set; } = System.Drawing.Color.Transparent;
        public ReoGridHorAlign HAlign { get; set; } = ReoGridHorAlign.General;
        public ReoGridVerAlign VAlign { get; set; } = ReoGridVerAlign.Middle;
    }
} 