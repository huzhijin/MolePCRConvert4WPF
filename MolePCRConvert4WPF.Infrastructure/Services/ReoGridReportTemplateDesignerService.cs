using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using MolePCRConvert4WPF.Core.Interfaces;
using MolePCRConvert4WPF.Core.Models;
using unvell.ReoGrid;
using MolePCRConvert4WPF.Infrastructure.Extensions;

namespace MolePCRConvert4WPF.Infrastructure.Services
{
    /// <summary>
    /// 基于ReoGrid的报告模板设计器服务
    /// </summary>
    public class ReoGridReportTemplateDesignerService : IReportTemplateDesignerService
    {
        private readonly ILogger<ReoGridReportTemplateDesignerService> _logger;
        private readonly string _templatesDirPath;

        public ReoGridReportTemplateDesignerService(ILogger<ReoGridReportTemplateDesignerService> logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _templatesDirPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates");
            
            // 确保模板目录存在
            if (!Directory.Exists(_templatesDirPath))
            {
                Directory.CreateDirectory(_templatesDirPath);
                _logger.LogInformation("创建模板目录: {TemplateDir}", _templatesDirPath);
            }
        }

        public async Task<ReportTemplate> CreateNewTemplateAsync(string templateName)
        {
            _logger.LogInformation("创建新模板: {TemplateName}", templateName);
            
            try
            {
                // 创建ReoGrid实例
                var workbook = new unvell.ReoGrid.ReoGridControl();
                
                // 添加默认工作表
                var worksheet = workbook.CurrentWorksheet;
                
                // 设置一些默认的单元格值
                worksheet["A1"] = "PCR分析报告模板";
                worksheet["A3"] = "板ID:";
                worksheet["B3"] = "[[PlateID]]";
                worksheet["A4"] = "板名称:";
                worksheet["B4"] = "[[PlateName]]";
                worksheet["A5"] = "日期:";
                worksheet["B5"] = "[[Date]]";
                
                worksheet["A7"] = "[[DataStart]]";
                
                // 保存模板
                string fileName = $"{templateName}.rgf";
                string filePath = Path.Combine(_templatesDirPath, fileName);
                
                // 如果文件已存在，添加时间戳
                if (File.Exists(filePath))
                {
                    fileName = $"{templateName}_{DateTime.Now:yyyyMMddHHmmss}.rgf";
                    filePath = Path.Combine(_templatesDirPath, fileName);
                }
                
                // 保存为ReoGrid文件格式
                using (var ms = new MemoryStream())
                {
                    workbook.Save(ms, unvell.ReoGrid.IO.FileFormat.ReoGridFormat);
                    byte[] gridData = ms.ToArray();
                    
                    await File.WriteAllBytesAsync(filePath, gridData);
                }
                
                // 创建模板对象
                var template = new ReportTemplate
                {
                    Id = Guid.NewGuid(),
                    Name = templateName,
                    FilePath = filePath,
                    IsExcelTemplate = false, // 这是ReoGrid模板
                    IsReoGridTemplate = true,
                    Description = "使用ReoGrid设计器创建的模板",
                    CreatedAt = DateTime.Now
                };
                
                _logger.LogInformation("新模板已创建: {TemplatePath}", filePath);
                return template;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "创建新模板时出错");
                throw;
            }
        }

        public async Task<ReportTemplate> SaveTemplateAsync(ReportTemplate template, byte[] gridData)
        {
            _logger.LogInformation("保存模板: {TemplateName}", template.Name);
            
            try
            {
                // 确保文件路径有效
                if (string.IsNullOrEmpty(template.FilePath))
                {
                    string fileName = $"{template.Name}.rgf";
                    template.FilePath = Path.Combine(_templatesDirPath, fileName);
                }
                
                // 保存ReoGrid数据
                await File.WriteAllBytesAsync(template.FilePath, gridData);
                
                // 更新模板
                template.UpdatedAt = DateTime.Now;
                template.IsReoGridTemplate = true;
                
                _logger.LogInformation("模板已保存: {TemplatePath}", template.FilePath);
                return template;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "保存模板时出错");
                throw;
            }
        }

        public async Task<byte[]> LoadTemplateAsync(ReportTemplate template)
        {
            _logger.LogInformation("加载模板: {TemplateName}", template.Name);
            
            try
            {
                if (!File.Exists(template.FilePath))
                {
                    throw new FileNotFoundException("模板文件不存在", template.FilePath);
                }
                
                // 读取ReoGrid文件数据
                byte[] gridData = await File.ReadAllBytesAsync(template.FilePath);
                
                _logger.LogInformation("模板已加载: {TemplatePath}", template.FilePath);
                return gridData;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "加载模板时出错");
                throw;
            }
        }

        public async Task<string> GenerateReportAsync(Plate plate, ReportTemplate template, string outputPath, string outputFileName = "", bool isPatientReport = false)
        {
            _logger.LogInformation("生成报告: Plate={PlateName}, Template={TemplateName}, IsPatientReport={IsPatientReport}", 
                plate.Name, template.Name, isPatientReport);
            
            try
            {
                if (!File.Exists(template.FilePath))
                {
                    throw new FileNotFoundException("模板文件不存在", template.FilePath);
                }
                
                // 设置输出文件名
                string reportFileName = string.IsNullOrEmpty(outputFileName) 
                    ? $"Report_{plate.Name}_{DateTime.Now:yyyyMMddHHmmss}.xlsx"
                    : outputFileName;
                    
                string reportFilePath = Path.Combine(outputPath, reportFileName);
                
                // 加载模板
                byte[] templateData = await File.ReadAllBytesAsync(template.FilePath);
                
                // 创建ReoGrid实例
                var workbook = new unvell.ReoGrid.ReoGridControl();
                
                // 加载模板数据
                using (var ms = new MemoryStream(templateData))
                {
                    workbook.Load(ms, unvell.ReoGrid.IO.FileFormat.ReoGridFormat);
                }
                
                var worksheet = workbook.CurrentWorksheet;
                
                // 替换通用占位符
                ReplaceTemplatePlaceholders(worksheet, plate);
                
                // 根据不同类型的报告，填充不同的数据
                if (isPatientReport)
                {
                    FillPatientReportData(worksheet, plate);
                }
                else
                {
                    FillPlateReportData(worksheet, plate);
                }
                
                // 保存为Excel文件
                workbook.Save(reportFilePath, unvell.ReoGrid.IO.FileFormat.Excel2007);
                
                _logger.LogInformation("报告已生成: {ReportPath}", reportFilePath);
                return reportFilePath;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "生成报告时出错");
                throw;
            }
        }

        public async Task<IEnumerable<ReportTemplate>> GetReportTemplatesAsync()
        {
            _logger.LogInformation("获取所有报告模板");
            
            try
            {
                var templates = new List<ReportTemplate>();
                
                if (!Directory.Exists(_templatesDirPath))
                {
                    Directory.CreateDirectory(_templatesDirPath);
                    _logger.LogInformation("创建模板目录: {TemplateDir}", _templatesDirPath);
                    return templates;
                }
                
                // 查找所有ReoGrid模板文件
                var rgfFiles = Directory.GetFiles(_templatesDirPath, "*.rgf");
                foreach (var file in rgfFiles)
                {
                    // 跳过以~$开头的临时文件
                    if (Path.GetFileName(file).StartsWith("~$")) continue;
                    
                    templates.Add(new ReportTemplate
                    {
                        Id = Guid.NewGuid(),
                        Name = Path.GetFileNameWithoutExtension(file),
                        FilePath = file,
                        IsExcelTemplate = false,
                        IsReoGridTemplate = true,
                        Description = "ReoGrid设计器模板",
                        CreatedAt = File.GetCreationTime(file)
                    });
                }
                
                // 查找所有Excel模板文件(向后兼容)
                var xlsxFiles = Directory.GetFiles(_templatesDirPath, "*.xlsx");
                foreach (var file in xlsxFiles)
                {
                    // 跳过以~$开头的临时文件
                    if (Path.GetFileName(file).StartsWith("~$")) continue;
                    
                    templates.Add(new ReportTemplate
                    {
                        Id = Guid.NewGuid(),
                        Name = Path.GetFileNameWithoutExtension(file),
                        FilePath = file,
                        IsExcelTemplate = true,
                        IsReoGridTemplate = false,
                        Description = "Excel模板(向后兼容)",
                        CreatedAt = File.GetCreationTime(file)
                    });
                }
                
                _logger.LogInformation("已加载 {Count} 个报告模板", templates.Count);
                return templates;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "获取报告模板时出错");
                throw;
            }
        }

        public async Task<bool> DeleteTemplateAsync(ReportTemplate template)
        {
            _logger.LogInformation("删除模板: {TemplateName}", template.Name);
            
            try
            {
                if (!File.Exists(template.FilePath))
                {
                    _logger.LogWarning("要删除的模板文件不存在: {TemplatePath}", template.FilePath);
                    return false;
                }
                
                File.Delete(template.FilePath);
                _logger.LogInformation("模板已删除: {TemplatePath}", template.FilePath);
                
                return true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "删除模板时出错");
                return false;
            }
        }

        #region 辅助方法

        /// <summary>
        /// 替换模板中的通用占位符
        /// </summary>
        private void ReplaceTemplatePlaceholders(unvell.ReoGrid.Worksheet worksheet, Plate plate)
        {
            // 定义占位符列表
            var placeholders = new Dictionary<string, string>
            {
                { "[[PlateID]]", plate.Id.ToString() },
                { "[[PlateName]]", plate.Name },
                { "[[Date]]", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") },
                { "[[TotalSamples]]", plate.WellLayouts.Count.ToString() },
                { "[[PositiveCount]]", plate.WellLayouts.Count(w => IsPositiveResult(w)).ToString() },
                { "[[NegativeCount]]", plate.WellLayouts.Count(w => IsNegativeResult(w)).ToString() }
            };
            
            // 遍历工作表所有单元格
            var range = worksheet.UsedRange;
            for (int row = range.Row; row < range.Row + range.Rows; row++)
            {
                for (int col = range.Col; col < range.Col + range.Cols; col++)
                {
                    var cell = worksheet.GetCell(row, col);
                    if (cell?.Data is string cellText && !string.IsNullOrEmpty(cellText))
                    {
                        string newText = cellText;
                        foreach (var placeholder in placeholders)
                        {
                            if (newText.Contains(placeholder.Key))
                            {
                                newText = newText.Replace(placeholder.Key, placeholder.Value);
                            }
                        }
                        
                        if (newText != cellText)
                        {
                            worksheet.SetCellData(row, col, newText);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 填充患者报告数据
        /// </summary>
        private void FillPatientReportData(unvell.ReoGrid.Worksheet worksheet, Plate plate)
        {
            // 按患者分组数据
            var patientGroups = plate.WellLayouts
                .Where(w => !string.IsNullOrEmpty(w.PatientName))
                .GroupBy(w => new { w.PatientName, w.PatientCaseNumber })
                .ToList();
            
            // 寻找数据表格起始位置标记 (例如 "[[DataStart]]")
            int? dataStartRow = null;
            int? dataStartCol = null;
            
            // 找到标记位置
            var range = worksheet.UsedRange;
            for (int row = range.Row; row < range.Row + range.Rows; row++)
            {
                for (int col = range.Col; col < range.Col + range.Cols; col++)
                {
                    var cell = worksheet.GetCell(row, col);
                    if (cell?.Data is string cellText && cellText == "[[DataStart]]")
                    {
                        dataStartRow = row;
                        dataStartCol = col;
                        // 清除标记
                        worksheet.SetCellData(row, col, "");
                        break;
                    }
                }
                if (dataStartRow.HasValue) break;
            }
            
            // 如果找到数据起始位置，填充数据
            if (dataStartRow.HasValue && dataStartCol.HasValue)
            {
                int currentRow = dataStartRow.Value;
                
                foreach (var patientGroup in patientGroups)
                {
                    var patientName = patientGroup.Key.PatientName;
                    var patientCaseNumber = patientGroup.Key.PatientCaseNumber;
                    
                    // 填充患者基本信息
                    worksheet.SetCellData(currentRow, dataStartCol.Value, patientName);
                    worksheet.SetCellData(currentRow, dataStartCol.Value + 1, patientCaseNumber);
                    
                    // 填充患者检测结果
                    foreach (var well in patientGroup)
                    {
                        worksheet.SetCellData(currentRow, dataStartCol.Value + 2, well.WellName);
                        worksheet.SetCellData(currentRow, dataStartCol.Value + 3, well.Channel);
                        worksheet.SetCellData(currentRow, dataStartCol.Value + 4, well.CtValue);
                        
                        // 获取结果文本
                        string resultText = GetResultText(well);
                        worksheet.SetCellData(currentRow, dataStartCol.Value + 5, resultText);
                        
                        // 如果是阳性结果，设置字体颜色为红色
                        if (IsPositiveResult(well))
                        {
                            var style = worksheet.GetRangeStyle(currentRow, dataStartCol.Value + 5, 1, 1);
                            style.TextColor = System.Drawing.Color.Red;
                            // 从我们的样式类转换为ReoGrid样式
                            var reoStyle = new unvell.ReoGrid.WorksheetRangeStyle
                            {
                                Bold = style.IsBold,
                                Italic = style.IsItalic,
                                Underline = style.IsUnderline,
                                TextColor = style.TextColor,
                                BackColor = style.BackColor,
                                HAlign = style.HAlign,
                                VAlign = style.VAlign
                            };
                            worksheet.SetRangeStyles(currentRow, dataStartCol.Value + 5, 1, 1, reoStyle);
                        }
                        
                        currentRow++;
                    }
                    
                    // 在患者之间添加空行
                    currentRow++;
                }
            }
        }

        /// <summary>
        /// 填充整板报告数据
        /// </summary>
        private void FillPlateReportData(unvell.ReoGrid.Worksheet worksheet, Plate plate)
        {
            // 寻找数据表格起始位置标记 (例如 "[[DataStart]]")
            int? dataStartRow = null;
            int? dataStartCol = null;
            
            // 找到标记位置
            var range = worksheet.UsedRange;
            for (int row = range.Row; row < range.Row + range.Rows; row++)
            {
                for (int col = range.Col; col < range.Col + range.Cols; col++)
                {
                    var cell = worksheet.GetCell(row, col);
                    if (cell?.Data is string cellText && cellText == "[[DataStart]]")
                    {
                        dataStartRow = row;
                        dataStartCol = col;
                        // 清除标记
                        worksheet.SetCellData(row, col, "");
                        break;
                    }
                }
                if (dataStartRow.HasValue) break;
            }
            
            // 如果找到了数据表格起始位置，填充数据
            if (dataStartRow.HasValue && dataStartCol.HasValue)
            {
                int currentRow = dataStartRow.Value;
                
                // 按孔位排序
                var sortedWells = plate.WellLayouts
                    .OrderBy(w => w.Row)
                    .ThenBy(w => w.Column)
                    .ToList();
                
                foreach (var well in sortedWells)
                {
                    // 填充基础数据
                    worksheet.SetCellData(currentRow, dataStartCol.Value, well.WellName);
                    worksheet.SetCellData(currentRow, dataStartCol.Value + 1, well.PatientName);
                    worksheet.SetCellData(currentRow, dataStartCol.Value + 2, well.PatientCaseNumber);
                    worksheet.SetCellData(currentRow, dataStartCol.Value + 3, well.Channel);
                    worksheet.SetCellData(currentRow, dataStartCol.Value + 4, well.CtValue);
                    
                    // 获取结果文本
                    string resultText = GetResultText(well);
                    worksheet.SetCellData(currentRow, dataStartCol.Value + 5, resultText);
                    
                    // 如果是阳性结果，设置字体颜色为红色
                    if (IsPositiveResult(well))
                    {
                        var style = worksheet.GetRangeStyle(currentRow, dataStartCol.Value + 5, 1, 1);
                        style.TextColor = System.Drawing.Color.Red;
                        // 从我们的样式类转换为ReoGrid样式
                        var reoStyle = new unvell.ReoGrid.WorksheetRangeStyle
                        {
                            Bold = style.IsBold,
                            Italic = style.IsItalic,
                            Underline = style.IsUnderline,
                            TextColor = style.TextColor,
                            BackColor = style.BackColor,
                            HAlign = style.HAlign,
                            VAlign = style.VAlign
                        };
                        worksheet.SetRangeStyles(currentRow, dataStartCol.Value + 5, 1, 1, reoStyle);
                    }
                    
                    currentRow++;
                }
            }
        }

        /// <summary>
        /// 判断孔位是否为阳性结果
        /// </summary>
        private bool IsPositiveResult(WellLayout well)
        {
            // 基于Ct值判断是否为阳性
            if (well.CtValue.HasValue && well.CtValue.Value < 35)
            {
                return true; // 如果Ct值小于35，认为是阳性
            }
            return false;
        }

        /// <summary>
        /// 判断孔位是否为阴性结果
        /// </summary>
        private bool IsNegativeResult(WellLayout well)
        {
            // 基于Ct值判断是否为阴性
            if (!well.CtValue.HasValue || well.CtValue.Value >= 35)
            {
                return true; // 如果Ct值未检出或>=35，认为是阴性
            }
            return false;
        }

        /// <summary>
        /// 获取孔位的结果文本描述
        /// </summary>
        private string GetResultText(WellLayout well)
        {
            if (IsPositiveResult(well))
            {
                return "阳性";
            }
            else if (IsNegativeResult(well))
            {
                return "阴性";
            }
            return "未知";
        }

        #endregion
    }
} 