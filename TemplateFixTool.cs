using System;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace MolePCRConvert4WPF
{
    public class TemplateFixTool
    {
        private static readonly Dictionary<string, string> PlaceholderMap = new Dictionary<string, string>
        {
            { "[受检人姓名]", "${PatientName}" },
            { "[受检人ID]", "${PatientId}" },
            { "[受检人性别]", "${Gender}" },
            { "[临床信息]", "${ClinicalInfo}" },
            { "[病历号]", "${PatientId}" },
            { "[样本编码]", "${SampleCode}" },
            { "[样本条码]", "${SampleId}" },
            { "[接收日期]", "${CollectionDate}" },
            { "[采样日期]", "${TestDate}" },
            { "[样本类型]", "${SampleType}" },
            { "[送检单位]", "${Department}" },
            { "[送检科室]", "${Section}" },
            { "[送检医生]", "${Doctor}" },
            { "[病原体名]", "${Target}" },
            { "[浓度值]", "${Concentration}" },
            { "[CT值]", "${CtValue}" },
            { "[检测结果]", "${Result}" },
            { "[备注]", "${Note}" },
            { "[实验编号]", "${ExperimentId}" }
        };

        /// <summary>
        /// 修复指定的Excel模板
        /// </summary>
        /// <param name="filePath">模板文件路径</param>
        /// <returns>修复结果</returns>
        public static (bool Success, string Message) FixTemplate(string filePath)
        {
            try
            {
                // 设置EPPlus许可证上下文
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                if (!File.Exists(filePath))
                {
                    return (false, $"文件不存在: {filePath}");
                }

                if (!filePath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                {
                    return (false, "只支持修复.xlsx格式的Excel文件");
                }

                // 创建备份
                string backupPath = $"{filePath}.backup";
                File.Copy(filePath, backupPath, true);

                // 打开Excel文件
                using (var package = new ExcelPackage(new FileInfo(filePath)))
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

                                    // 替换已知占位符
                                    foreach (var placeholder in PlaceholderMap)
                                    {
                                        if (newValue.Contains(placeholder.Key))
                                        {
                                            newValue = newValue.Replace(placeholder.Key, placeholder.Value);
                                            hasChanges = true;
                                        }
                                    }

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

                        // 2. 尝试寻找并添加数据表格标记
                        if (!hasDataStart)
                        {
                            // 查找表格头部，这通常包含"病原体名"、"CT值"等列名
                            for (int row = 1; row <= worksheet.Dimension.End.Row; row++)
                            {
                                bool foundHeader = false;
                                int headerCol = -1;

                                // 查找表头行
                                for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                                {
                                    string cellValue = worksheet.Cells[row, col].Text;
                                    if (cellValue != null && (
                                        cellValue.Contains("病原体") || 
                                        cellValue.Contains("CT值") || 
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
                                    break;
                                }
                            }
                        }
                    }

                    // 保存更改
                    if (hasChanges)
                    {
                        package.Save();
                        string message = "模板已修复：\n";
                        message += "- 替换了旧格式占位符为新格式\n";
                        if (hasDataStart)
                            message += "- 已添加或确认存在[[DataStart]]标记";
                        else
                            message += "- 警告：未能自动添加[[DataStart]]标记，请手动添加";

                        return (true, message);
                    }
                    else
                    {
                        return (true, "模板已检查，未发现需要修复的问题。");
                    }
                }
            }
            catch (Exception ex)
            {
                return (false, $"修复模板时出错：{ex.Message}");
            }
        }

        /// <summary>
        /// 修复指定目录下的所有Excel模板
        /// </summary>
        /// <param name="directoryPath">模板目录路径</param>
        /// <returns>修复结果列表</returns>
        public static List<(string FileName, bool Success, string Message)> FixAllTemplates(string directoryPath)
        {
            var results = new List<(string, bool, string)>();

            if (!Directory.Exists(directoryPath))
            {
                results.Add(("", false, $"目录不存在: {directoryPath}"));
                return results;
            }

            var files = Directory.GetFiles(directoryPath, "*.xlsx");
            foreach (var file in files)
            {
                // 跳过临时文件
                if (Path.GetFileName(file).StartsWith("~$")) continue;

                var (success, message) = FixTemplate(file);
                results.Add((Path.GetFileName(file), success, message));
            }

            return results;
        }
    }
} 