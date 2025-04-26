using System;
using System.Windows.Controls;
using System.Windows;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using OfficeOpenXml.Table;
using NPOI.SS.UserModel;
using System.Globalization;
using Microsoft.Extensions.Logging;
using MolePCRConvert4WPF.Core.Models;
using NPOI.SS.UserModel.Charts;

namespace MolePCRConvert4WPF.App.Views
{
    /// <summary>
    /// ReportTemplateConfigView.xaml 的交互逻辑
    /// </summary>
    public partial class ReportTemplateConfigView : UserControl
    {
        private readonly ILogger<ReportTemplateConfigView>? _logger;
        
        public ReportTemplateConfigView(ILogger<ReportTemplateConfigView>? logger = null)
        {
            _logger = logger;
            InitializeComponent();
        }

        /// <summary>
        /// 修复模板按钮点击事件
        /// </summary>
        private void BtnFixTemplates_Click(object sender, RoutedEventArgs e)
        {
            // 先询问用户是否强制修复
            var result = MessageBox.Show(
                "是否要强制修复所有模板？\n\n" +
                "- 选择\"是\"将强制对所有模板添加[[DataStart]]标记和替换占位符\n" +
                "- 选择\"否\"将仅修复有明显问题的模板\n\n" +
                "注意：修复前将自动创建备份文件(.backup后缀)",
                "模板修复选项",
                MessageBoxButton.YesNoCancel,
                MessageBoxImage.Question);

            if (result == MessageBoxResult.Cancel)
                return;

            bool forceFixAll = (result == MessageBoxResult.Yes);

            // 获取模板目录
            string templateDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates");
            if (!Directory.Exists(templateDir))
            {
                MessageBox.Show($"模板目录不存在: {templateDir}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            // 调用模板修复方法
            bool anyFixed = false;
            string message = "模板检查结果:\n";
            
            var files = Directory.GetFiles(templateDir, "*.xlsx");
            foreach (var file in files)
            {
                // 跳过临时文件
                if (Path.GetFileName(file).StartsWith("~$")) continue;
                
                try
                {
                    // 设置EPPlus许可证上下文
                    OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                    
                    // 创建备份
                    string backupPath = $"{file}.backup";
                    File.Copy(file, backupPath, true);
                    
                    // 打开Excel文件
                    using (var package = new OfficeOpenXml.ExcelPackage(new FileInfo(file)))
                    {
                        bool hasChanges = false;
                        bool hasDataStart = false;
                        
                        foreach (var worksheet in package.Workbook.Worksheets)
                        {
                            if (worksheet.Dimension == null) continue;
                            
                            // 1. 替换占位符
                            for (int row = 1; row <= worksheet.Dimension.End.Row; row++)
                            {
                                for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                                {
                                    var cell = worksheet.Cells[row, col];
                                    string cellValue = cell.Text;
                                    
                                    if (!string.IsNullOrEmpty(cellValue))
                                    {
                                        string newValue = cellValue;
                                        
                                        // 替换所有已知占位符 - 扩展了占位符列表
                                        if (newValue.Contains("[受检人姓名]")) { newValue = newValue.Replace("[受检人姓名]", "${PatientName}"); hasChanges = true; }
                                        if (newValue.Contains("[受检人ID]")) { newValue = newValue.Replace("[受检人ID]", "${PatientId}"); hasChanges = true; }
                                        if (newValue.Contains("[受检人性别]")) { newValue = newValue.Replace("[受检人性别]", "${Gender}"); hasChanges = true; }
                                        if (newValue.Contains("[病历号]")) { newValue = newValue.Replace("[病历号]", "${PatientId}"); hasChanges = true; }
                                        if (newValue.Contains("[样本条码]")) { newValue = newValue.Replace("[样本条码]", "${SampleId}"); hasChanges = true; }
                                        if (newValue.Contains("[样本编码]")) { newValue = newValue.Replace("[样本编码]", "${SampleCode}"); hasChanges = true; }
                                        if (newValue.Contains("[接收日期]")) { newValue = newValue.Replace("[接收日期]", "${CollectionDate}"); hasChanges = true; }
                                        if (newValue.Contains("[采样日期]")) { newValue = newValue.Replace("[采样日期]", "${TestDate}"); hasChanges = true; }
                                        if (newValue.Contains("[样本类型]")) { newValue = newValue.Replace("[样本类型]", "${SampleType}"); hasChanges = true; }
                                        if (newValue.Contains("[送检单位]")) { newValue = newValue.Replace("[送检单位]", "${Department}"); hasChanges = true; }
                                        if (newValue.Contains("[送检科室]")) { newValue = newValue.Replace("[送检科室]", "${Section}"); hasChanges = true; }
                                        if (newValue.Contains("[送检医生]")) { newValue = newValue.Replace("[送检医生]", "${Doctor}"); hasChanges = true; }
                                        if (newValue.Contains("[病原体名]")) { newValue = newValue.Replace("[病原体名]", "${Target}"); hasChanges = true; }
                                        if (newValue.Contains("[浓度值]")) { newValue = newValue.Replace("[浓度值]", "${Concentration}"); hasChanges = true; }
                                        if (newValue.Contains("[CT值]")) { newValue = newValue.Replace("[CT值]", "${CtValue}"); hasChanges = true; }
                                        if (newValue.Contains("[检测结果]")) { newValue = newValue.Replace("[检测结果]", "${Result}"); hasChanges = true; }
                                        if (newValue.Contains("[备注]")) { newValue = newValue.Replace("[备注]", "${Note}"); hasChanges = true; }
                                        
                                        // 检查是否已有[[DataStart]]标记
                                        if (cellValue.Contains("[[DataStart]]"))
                                        {
                                            hasDataStart = true;
                                        }
                                        
                                        // 更新单元格值
                                        if (newValue != cellValue)
                                        {
                                            cell.Value = newValue;
                                        }
                                    }
                                }
                            }
                            
                            // 2. 检查并添加数据表格标记
                            // 如果强制修复或未找到[[DataStart]]标记，尝试添加
                            if (forceFixAll || !hasDataStart)
                            {
                                bool addedDataStart = false;
                                
                                // A. 首先尝试查找表格头部
                                for (int row = 1; row <= worksheet.Dimension.End.Row; row++)
                                {
                                    bool foundHeader = false;
                                    int headerCol = -1;
                                    
                                    // 查找常见表头行
                                    for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                                    {
                                        string cellValue = worksheet.Cells[row, col].Text;
                                        if (cellValue != null && (
                                            cellValue.Contains("病原体") || 
                                            cellValue.Contains("CT值") || 
                                            cellValue.Contains("靶标") ||
                                            cellValue.Contains("结果")))
                                        {
                                            foundHeader = true;
                                            headerCol = col;
                                            break;
                                        }
                                    }
                                    
                                    // 在表头下方的首个单元格添加[[DataStart]]标记
                                    if (foundHeader && row < worksheet.Dimension.End.Row)
                                    {
                                        worksheet.Cells[row + 1, headerCol].Value = "[[DataStart]]";
                                        hasDataStart = true;
                                        hasChanges = true;
                                        addedDataStart = true;
                                        break;
                                    }
                                }
                                
                                // B. 如果找不到表头，则寻找可能的数据区域
                                if (!addedDataStart && forceFixAll)
                                {
                                    // 从第10行开始往下查找空行后的第一个非空单元格
                                    for (int row = 10; row <= Math.Min(worksheet.Dimension.End.Row, 30); row++)
                                    {
                                        bool isEmptyRow = true;
                                        for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                                        {
                                            if (!string.IsNullOrEmpty(worksheet.Cells[row, col].Text))
                                            {
                                                isEmptyRow = false;
                                                break;
                                            }
                                        }
                                        
                                        // 如果找到空行，检查下一行
                                        if (isEmptyRow && row < worksheet.Dimension.End.Row)
                                        {
                                            // 在下一行第一个非空单元格添加标记
                                            for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                                            {
                                                if (!string.IsNullOrEmpty(worksheet.Cells[row + 1, col].Text))
                                                {
                                                    worksheet.Cells[row + 1, col].Value = "[[DataStart]]";
                                                    hasDataStart = true;
                                                    hasChanges = true;
                                                    addedDataStart = true;
                                                    break;
                                                }
                                            }
                                            
                                            if (addedDataStart) break;
                                        }
                                    }
                                }
                                
                                // C. 最后尝试，如果还是找不到合适位置，强制在第15行添加
                                if (!addedDataStart && forceFixAll)
                                {
                                    int targetRow = Math.Min(15, worksheet.Dimension.End.Row);
                                    worksheet.Cells[targetRow, 1].Value = "[[DataStart]]";
                                    hasDataStart = true;
                                    hasChanges = true;
                                }
                            }
                        }
                        
                        // 保存更改
                        if (hasChanges)
                        {
                            package.Save();
                            message += $"- {Path.GetFileName(file)}: 已修复\n";
                            anyFixed = true;
                        }
                        else
                        {
                            if (forceFixAll)
                            {
                                // 强制添加标记，即使没有找到需要替换的占位符
                                var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                                if (worksheet != null && worksheet.Dimension != null)
                                {
                                    worksheet.Cells[15, 1].Value = "[[DataStart]]";
                                    package.Save();
                                    message += $"- {Path.GetFileName(file)}: 已强制添加数据标记\n";
                                    anyFixed = true;
                                }
                                else
                                {
                                    message += $"- {Path.GetFileName(file)}: 无法修复 (工作表为空)\n";
                                }
                            }
                            else
                            {
                                message += $"- {Path.GetFileName(file)}: 无需修复\n";
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    message += $"- {Path.GetFileName(file)}: 处理出错 - {ex.Message}\n";
                }
            }
            
            if (anyFixed)
            {
                MessageBox.Show(message, "模板修复完成", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                MessageBox.Show(message, "模板检查完成", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }

        private void ParseSLANData(ISheet sheet, int headerRow, int wellPositionColumn, int targetColumn, int ctColumn)
        {
            // 确保从第15行开始读取(表头是第14行)
            for (int rowIndex = headerRow + 1; rowIndex <= sheet.LastRowNum; rowIndex++)
            {
                var row = sheet.GetRow(rowIndex);
                if (row == null) continue;

                // 从A列读取孔位信息(例如"A1")
                var wellPositionCell = row.GetCell(wellPositionColumn);
                if (wellPositionCell == null) continue;
                
                string wellPosition = wellPositionCell.StringCellValue?.Trim();
                if (string.IsNullOrEmpty(wellPosition)) continue;
                
                // SLAN特定：检查孔位格式并确保其符合应用程序的期望格式
                // 例如，确保它是类似"A1"的格式
                if (!Regex.IsMatch(wellPosition, @"^[A-H][1-9][0-2]?$"))
                {
                    _logger?.LogWarning($"无法解析孔位名称格式: '{wellPosition}' (行索引 {rowIndex})");
                    continue;
                }
                
                // 从G列(通道列)读取通道信息
                var channelCell = row.GetCell(6); // G列索引为6
                string channel = channelCell?.StringCellValue?.Trim();
                if (string.IsNullOrEmpty(channel)) continue;
                
                // 从P列读取CT值
                var ctCell = row.GetCell(ctColumn);
                double? ctValue = null;
                if (ctCell != null)
                {
                    if (ctCell.CellType == CellType.Numeric)
                    {
                        ctValue = ctCell.NumericCellValue;
                    }
                    else if (ctCell.CellType == CellType.String)
                    {
                        if (double.TryParse(ctCell.StringCellValue, out double parsedValue))
                        {
                            ctValue = parsedValue;
                        }
                    }
                }
                
                // 从H列读取目标信息
                var targetCell = row.GetCell(targetColumn);
                string target = targetCell?.StringCellValue?.Trim() ?? string.Empty;
                
                // 创建WellLayout对象并添加到结果中
                var well = new WellLayout
                {
                    Row = wellPosition.Substring(0, 1),
                    Column = int.Parse(wellPosition.Substring(1)),
                    Channel = target,
                    CtValue = ctValue,
                    TargetName = target
                };
                
                // 添加到处理结果中
            }
        }
    }
} 