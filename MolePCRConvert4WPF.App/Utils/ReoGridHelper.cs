using System.Drawing;
using unvell.ReoGrid;
using unvell.ReoGrid.Graphics;
using MolePCRConvert4WPF.App.Commands;
using CommunityToolkit.Mvvm.Input;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Input;
using MolePCRConvert4WPF.App.ViewModels;
using MolePCRConvert4WPF.Core.Models;
using unvell.ReoGrid.Actions;

namespace MolePCRConvert4WPF.App.Utils
{
    /// <summary>
    /// ReoGrid辅助类，用于处理ReoGrid相关的转换和扩展功能
    /// </summary>
    public static class ReoGridHelper
    {
        /// <summary>
        /// 表示ReoGrid的位置
        /// </summary>
        public struct ReoGridPos
        {
            public int Row { get; set; }
            public int Col { get; set; }

            public ReoGridPos(int row, int col)
            {
                Row = row;
                Col = col;
            }

            public static implicit operator CellPosition(ReoGridPos pos)
            {
                return new CellPosition(pos.Row, pos.Col);
            }

            public static implicit operator ReoGridPos(CellPosition pos)
            {
                return new ReoGridPos(pos.Row, pos.Col);
            }
            
            public override string ToString()
            {
                string colStr = RGUtility.GetAlphaChar(Col);
                return $"{colStr}{Row + 1}";
            }
        }

        /// <summary>
        /// 将System.Drawing.Color转换为unvell.ReoGrid.Graphics.SolidColor
        /// </summary>
        public static SolidColor ToSolidColor(this Color color)
        {
            return SolidColor.FromArgb((byte)color.A, (byte)color.R, (byte)color.G, (byte)color.B);
        }

        /// <summary>
        /// 将unvell.ReoGrid.Graphics.SolidColor转换为System.Drawing.Color
        /// </summary>
        public static Color ToDrawingColor(this SolidColor color)
        {
            return Color.FromArgb(color.A, color.R, color.G, color.B);
        }

        /// <summary>
        /// 设置工作表行数
        /// </summary>
        public static void SetRowsCount(this Worksheet worksheet, int count)
        {
            worksheet.SetRows(count);
        }

        /// <summary>
        /// 设置工作表列数
        /// </summary>
        public static void SetColumnsCount(this Worksheet worksheet, int count)
        {
            worksheet.SetCols(count);
        }

        /// <summary>
        /// 设置工作表默认行高
        /// </summary>
        public static void SetDefaultRowHeight(this Worksheet worksheet, int height)
        {
            worksheet.SetRowsHeight(0, worksheet.RowCount, (ushort)height);
        }

        /// <summary>
        /// 设置工作表默认列宽
        /// </summary>
        public static void SetDefaultColumnWidth(this Worksheet worksheet, int width)
        {
            worksheet.SetColumnsWidth(0, worksheet.ColumnCount, (ushort)width);
        }

        /// <summary>
        /// 设置工作表默认行/列大小
        /// </summary>
        public static void SetDefaults(this Worksheet worksheet, RowOrColumn rowOrColumn, int size)
        {
            if (rowOrColumn == RowOrColumn.Row)
            {
                SetDefaultRowHeight(worksheet, size);
            }
            else
            {
                SetDefaultColumnWidth(worksheet, size);
            }
        }

        /// <summary>
        /// 获取ReoGrid单元格样式的粗体设置
        /// </summary>
        public static bool GetBold(this WorksheetRangeStyle style)
        {
            return style.Bold;
        }

        /// <summary>
        /// 设置ReoGrid单元格样式的粗体
        /// </summary>
        public static void SetBold(this WorksheetRangeStyle style, bool value)
        {
            style.Bold = value;
        }
    }
    
    /// <summary>
    /// 命令相关扩展方法
    /// </summary>
    public static class CommandExtensions
    {
        /// <summary>
        /// 通知命令可执行状态已改变的扩展方法
        /// </summary>
        public static void RaiseCanExecuteChanged(this MolePCRConvert4WPF.App.Commands.RelayCommand command)
        {
            if (command != null)
            {
                command.RaiseCanExecuteChanged();
            }
        }

        /// <summary>
        /// 通知命令可执行状态已改变的扩展方法
        /// </summary>
        public static void RaiseCanExecuteChanged<T>(this MolePCRConvert4WPF.App.Commands.RelayCommand<T> command)
        {
            if (command != null)
            {
                command.RaiseCanExecuteChanged();
            }
        }
        
        /// <summary>
        /// 通知CommunityToolkit命令可执行状态已改变的扩展方法
        /// </summary>
        public static void NotifyCanExecuteChangedIfNeeded(this System.Windows.Input.ICommand command)
        {
            if (command is CommunityToolkit.Mvvm.Input.IRelayCommand relayCommand)
            {
                relayCommand.NotifyCanExecuteChanged();
            }
            else if (command is MolePCRConvert4WPF.App.Commands.RelayCommand legacyCommand)
            {
                legacyCommand.RaiseCanExecuteChanged();
            }
            else if (command is MolePCRConvert4WPF.App.ViewModels.RelayCommand oldCommand)
            {
                // 处理旧的RelayCommand类型
                try
                {
                    // 使用反射调用RaiseCanExecuteChanged方法
                    var method = oldCommand.GetType().GetMethod("RaiseCanExecuteChanged");
                    if (method != null)
                    {
                        method.Invoke(oldCommand, null);
                    }
                    else
                    {
                        // 如果没有找到方法，回退到全局更新
                        System.Windows.Input.CommandManager.InvalidateRequerySuggested();
                    }
                }
                catch
                {
                    // 最后的回退机制：强制触发CanExecuteChanged事件
                    System.Windows.Input.CommandManager.InvalidateRequerySuggested();
                }
            }
            else
            {
                // 对于其他类型的命令，直接使用全局刷新机制
                System.Windows.Input.CommandManager.InvalidateRequerySuggested();
            }
        }
    }
} 