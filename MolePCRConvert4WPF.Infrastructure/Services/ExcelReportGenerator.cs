using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using MolePCRConvert4WPF.Core.Interfaces;
using MolePCRConvert4WPF.Core.Models;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;
using System.Text;

namespace MolePCRConvert4WPF.Infrastructure.Services
{
    /// <summary>
    /// 使用 EPPlus 生成 Excel 报告的服务
    /// </summary>
    public class ExcelReportGenerator : IReportService
    {
        private readonly ILogger<ExcelReportGenerator> _logger;

        // 静态初始化 EPPlus 许可证上下文
        static ExcelReportGenerator()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        public ExcelReportGenerator(ILogger<ExcelReportGenerator> logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        /// <summary>
        /// 兼容旧版接口的方法
        /// </summary>
        public async Task<string> GenerateExcelReportAsync(Plate plate, ReportTemplate template, string outputPath)
        {
            // 调用新的接口方法 - 注意：旧接口缺少 analysisResults，这里需要处理
            // 方案1: 抛出异常，强制使用新接口
            // throw new NotSupportedException("旧版 GenerateExcelReportAsync 接口已弃用，请使用包含 analysisResults 的新接口。");
            // 方案2: 尝试获取 analysisResults (如果可能的话，但这违反了依赖注入原则)
            // 方案3: 传递一个空的列表，报告内容会不完整
             _logger.LogWarning("正在调用旧版 GenerateExcelReportAsync 接口，无法获取分析结果，报告内容可能不完整。");
            return await GenerateExcelReportAsync(plate, template, outputPath, new List<AnalysisResultItem>(), "", false);
        }

        /// <summary>
        /// 生成Excel报告 (使用分析结果)
        /// </summary>
        public async Task<string> GenerateExcelReportAsync(Plate plate, ReportTemplate template, string outputPath, IEnumerable<AnalysisResultItem> analysisResults, string outputFileName = "", bool isPatientReport = false)
        {
            _logger.LogInformation("开始生成 Excel 报告: Plate={PlateName}, Template={TemplateName}, IsPatientReport={IsPatientReport}", 
                plate.Name, template.Name, isPatientReport);

            string reportFileName = string.IsNullOrEmpty(outputFileName) 
                ? $"Report_{plate.Name}_{DateTime.Now:yyyyMMddHHmmss}.xlsx"
                : outputFileName;
            string reportFilePath = Path.Combine(outputPath, reportFileName);

            try
            {
                FileInfo templateFile = new FileInfo(template.FilePath);
                FileInfo reportFile = new FileInfo(reportFilePath);

                if (!templateFile.Exists)
                {
                    _logger.LogError("报告模板文件不存在: {TemplatePath}", template.FilePath);
                    throw new FileNotFoundException("报告模板文件未找到", template.FilePath);
                }

                using (var package = new ExcelPackage(reportFile, templateFile))
                {
                    var worksheet = package.Workbook.Worksheets.FirstOrDefault(); 
                    if (worksheet != null)
                    {                        
                        _logger.LogInformation("正在填充报告数据到工作表: {WorksheetName}", worksheet.Name);
                        await ReplaceTemplatePlaceholders(worksheet, plate);
                        
                        if (isPatientReport)
                        {
                            // 获取第一个患者的信息用于填充模板头部占位符（如果需要）
                            var firstPatientResult = analysisResults.FirstOrDefault();
                            string patientNameForHeader = firstPatientResult?.PatientName ?? plate.WellLayouts.FirstOrDefault()?.PatientName ?? "未知患者";
                            string caseNumberForHeader = firstPatientResult?.PatientCaseNumber ?? plate.WellLayouts.FirstOrDefault()?.PatientCaseNumber ?? "";
                            
                            ReplacePatientPlaceholders(worksheet, patientNameForHeader, caseNumberForHeader);
                            
                            // 处理病原体列表，传入分析结果
                            await FindAndProcessPatientResults(worksheet, analysisResults); // 使用分析结果
                        }
                        else
                        {
                            // 整板报告逻辑可能也需要调整以使用 analysisResults，暂时保留旧逻辑
                            await FillPlateReportData(worksheet, plate); 
                        }
                    }
                    else
                    {
                         _logger.LogWarning("在模板中找不到用于填充数据的工作表");
                    }
                    
                    await package.SaveAsync();
                    _logger.LogInformation("Excel 报告已保存到: {ReportPath}", reportFilePath);
                }

                return reportFilePath;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "生成 Excel 报告时出错");
                throw; 
            }
        }

        /// <summary>
        /// 生成报告预览 (使用分析结果)
        /// </summary>
        public async Task<(List<string> HtmlPreviews, List<string> PatientNames)> GenerateReportPreviewAsync(Plate plate, ReportTemplate template, IEnumerable<AnalysisResultItem> analysisResults, bool isPatientReport = false)
        {
            _logger.LogInformation("开始生成报告预览: Plate={PlateName}, Template={TemplateName}, IsPatientReport={IsPatientReport}", 
                plate.Name, template.Name, isPatientReport);
            
            List<string> htmlPreviews = new List<string>();
            List<string> patientNames = new List<string>();
            
            try
            {
                // 使用Try-Catch包裹所有可能出现问题的代码块
                try
                {
                    FileInfo templateFile = new FileInfo(template.FilePath);
                    
                    if (!templateFile.Exists)
                    {
                        _logger.LogError("报告模板文件不存在: {TemplatePath}", template.FilePath);
                        throw new FileNotFoundException("报告模板文件未找到", template.FilePath);
                    }
                    
                    // 尝试自动修复模板中的常见问题
                    await AutoFixTemplateIfNeededAsync(template);
                    
                    // 创建临时文件路径用于预览
                    string tempDir = Path.Combine(Path.GetTempPath(), "ReportPreviews");
                    if (!Directory.Exists(tempDir))
                        Directory.CreateDirectory(tempDir);
                    
                    string tempFilePath = Path.Combine(tempDir, $"Preview_{Guid.NewGuid()}.xlsx");
                    
                    // 使用临时文件生成Excel
                    using (var package = new ExcelPackage(new FileInfo(tempFilePath), templateFile))
                    {
                        var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                        if (worksheet != null)
                        {
                            // 检查工作表是否有内容
                            if (worksheet.Dimension == null)
                            {
                                _logger.LogWarning("报告模板工作表为空");
                                throw new InvalidOperationException("报告模板工作表为空，无法生成预览");
                            }

                            // 替换占位符
                            ReplaceTemplatePlaceholders(worksheet, plate);
                            
                            if (isPatientReport)
                            {
                                // 按患者分组分析结果
                                var patientResultGroups = analysisResults
                                    .Where(r => !string.IsNullOrEmpty(r.PatientName))
                                    .GroupBy(r => new { r.PatientName, r.PatientCaseNumber })
                                    .ToList();
                                
                                foreach (var patientGroup in patientResultGroups)
                                {
                                    // 创建临时工作表副本
                                    var patientWorksheet = package.Workbook.Worksheets.Copy(worksheet.Name, $"{worksheet.Name}_{htmlPreviews.Count}");
                                    
                                    // 替换患者信息占位符
                                    ReplacePatientPlaceholders(patientWorksheet, patientGroup.Key.PatientName, patientGroup.Key.PatientCaseNumber);
                                    
                                    // 处理病原体列表，传入当前患者的分析结果
                                    await FindAndProcessPatientResults(patientWorksheet, patientGroup.ToList()); // 使用当前患者的分析结果
                                    
                                    // 将Excel转为HTML
                                    string html = ConvertWorksheetToHtml(patientWorksheet);
                                    htmlPreviews.Add(html);
                                    
                                    // 添加患者名称
                                    patientNames.Add(patientGroup.Key.PatientName);
                                    
                                    // 从包中移除临时工作表以避免内存泄漏
                                    package.Workbook.Worksheets.Delete(patientWorksheet);
                                }
                            }
                            else
                            {
                                // 整板报告预览逻辑 - 注意：这里可能仍需调整以正确显示所有患者的分析结果
                                // 暂时保留旧逻辑，但 FindAndProcessPatientResults 内部逻辑已改变
                                await FillPlateReportData(worksheet, plate); // 旧方法，可能需要更新
                                // 或者，如果整板报告也需要展示分析结果的聚合/列表，需要类似患者报告的逻辑
                                // await FindAndProcessPatientResults(worksheet, analysisResults); // 使用所有分析结果填充
                                
                                string html = ConvertWorksheetToHtml(worksheet);
                                // ... (省略构建整板报告HTML的代码, 与之前版本类似，但可能需要调整以反映使用了 analysisResults) ...
                                htmlPreviews.Add(html); // 简化处理，直接添加转换后的HTML
                                patientNames.Add("整板报告");
                            }
                        }
                        else
                        {
                            _logger.LogWarning("在模板中找不到工作表");
                            throw new InvalidOperationException("模板文件中未找到有效的工作表");
                        }
                        
                        // 保存临时文件以便调试
                        await package.SaveAsync();
                        _logger.LogDebug("临时预览Excel文件已保存: {TempPath}", tempFilePath);
                    }
                }
                catch (ArgumentOutOfRangeException ex) when (ex.Message.Contains("startIndex cannot be larger"))
                {
                    // 特别处理字符串索引越界异常
                    _logger.LogError(ex, "生成HTML预览时发生字符串索引越界错误");
                    htmlPreviews.Add("生成表格预览时发生错误: startIndex cannot be larger than length of string. (Parameter 'startIndex')");
                    patientNames.Add("错误");
                    return (htmlPreviews, patientNames);
                }
                catch (Exception ex)
                {
                    // 捕获其他所有异常
                    _logger.LogError(ex, "生成Excel预览过程中出错");
                    throw; // 重新抛出以便被外层catch处理
                }
                
                // 如果没有生成任何预览，添加一个提示消息
                if (htmlPreviews.Count == 0)
                {
                    htmlPreviews.Add("<html><body><h2 style='color:#666;text-align:center;margin-top:100px;'>没有找到可供预览的内容</h2><p style='text-align:center;'>请检查模板格式是否包含必要的表格或标记</p></body></html>");
                    patientNames.Add("无数据");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "生成报告预览时出错，异常类型: {ExType}", ex.GetType().Name);
                
                // 创建更详细的错误信息HTML页面
                StringBuilder errorHtml = new StringBuilder();
                errorHtml.AppendLine("<!DOCTYPE html><html><head><meta charset='utf-8'>");
                errorHtml.AppendLine("<style>body{font-family:Arial;margin:20px;} .error{color:red;} .detail{color:#666;margin:10px 0;}</style>");
                errorHtml.AppendLine("</head><body>");
                errorHtml.AppendLine("<h2 class='error'>生成预览时发生错误</h2>");
                errorHtml.AppendLine($"<p class='error'>{ex.Message}</p>");
                errorHtml.AppendLine("<p class='detail'>错误可能原因：</p>");
                errorHtml.AppendLine("<ul class='detail'>");
                errorHtml.AppendLine("<li>模板格式不兼容或缺少必要的占位符</li>");
                errorHtml.AppendLine("<li>模板中的表格结构有问题</li>");
                errorHtml.AppendLine("<li>系统无法处理模板中的某些元素</li>");
                errorHtml.AppendLine("</ul>");
                errorHtml.AppendLine("<p class='detail'>建议：</p>");
                errorHtml.AppendLine("<ul class='detail'>");
                errorHtml.AppendLine("<li>检查模板文件是否包含正确的占位符（使用${xxx}格式）</li>");
                errorHtml.AppendLine("<li>确保在数据表格开始处添加了[[DataStart]]标记</li>");
                errorHtml.AppendLine("<li>尝试使用更简单的模板格式</li>");
                errorHtml.AppendLine("<li>尝试直接导出Excel而不预览</li>");
                errorHtml.AppendLine("</ul>");
                errorHtml.AppendLine("</body></html>");
                
                htmlPreviews.Add(errorHtml.ToString());
                patientNames.Add("错误");
            }
            
            return (htmlPreviews, patientNames);
        }
        
        /// <summary>
        /// 清除工作表数据，保留格式
        /// </summary>
        private void ClearWorksheetData(ExcelWorksheet worksheet)
        {
            // 清除所有单元格值，但保留样式和格式
            for (int i = 1; i <= worksheet.Dimension.End.Row; i++)
            {
                for (int j = 1; j <= worksheet.Dimension.End.Column; j++)
                {
                    // 清除值但保留样式
                    var value = worksheet.Cells[i, j].Value;
                    if (value != null && !IsTableHeader(worksheet.Cells[i, j].Text))
                    {
                        worksheet.Cells[i, j].Value = null;
                    }
                }
            }
        }
        
        /// <summary>
        /// 检查单元格是否是表头
        /// </summary>
        private bool IsTableHeader(string cellText)
        {
            // 常见的表头文本
            string[] headerTexts = new string[] 
            { 
                "孔位", "患者姓名", "病历号", "通道", "Ct值", "结果", 
                "位置", "样本ID", "患者ID", "靶标", "浓度" 
            };
            
            return headerTexts.Any(h => cellText.Contains(h));
        }
        
        /// <summary>
        /// 将工作表转换为HTML预览
        /// </summary>
        private string ConvertWorksheetToHtml(ExcelWorksheet worksheet)
        {
            StringBuilder html = new StringBuilder();
            
            html.AppendLine("<!DOCTYPE html>");
            html.AppendLine("<html>");
            html.AppendLine("<head>");
            html.AppendLine("<meta charset='utf-8'>");
            html.AppendLine("<style>");
            html.AppendLine("body { font-family: Arial, sans-serif; margin: 20px; }");
            html.AppendLine("table { border-collapse: collapse; width: 100%; }");
            html.AppendLine("th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }");
            html.AppendLine("th { background-color: #f2f2f2; }");
            html.AppendLine("tr:nth-child(even) { background-color: #f9f9f9; }");
            html.AppendLine(".positive { color: red; font-weight: bold; }");
            html.AppendLine(".negative { color: green; }");
            html.AppendLine(".report-content { position: relative; }");
            html.AppendLine(".report-background { position: absolute; top: 0; left: 0; width: 100%; height: 100%; z-index: -1; opacity: 0.3; }");
            html.AppendLine("</style>");
            html.AppendLine("</head>");
            html.AppendLine("<body>");
            
            try
            {
                // 检查worksheet和dimension是否有效
                if (worksheet == null || worksheet.Dimension == null)
                {
                    _logger.LogWarning("无效的工作表或工作表为空");
                    html.AppendLine("<p style='color:red;'>错误：无法读取工作表数据</p>");
                    html.AppendLine("</body>");
                    html.AppendLine("</html>");
                    return html.ToString();
                }
                
                // 检查工作表中的图片
                var drawings = worksheet.Drawings;
                if (drawings.Count > 0)
                {
                    html.AppendLine("<div class='report-content'>");
                    
                    // 注意：这里仅提示有图片，但HTML预览中无法显示Excel中的图片
                    // 未来可以考虑将图片导出为临时文件并链接
                    html.AppendLine("<p style='color:#666;text-align:center;'>(报表包含图片背景，在Excel中查看完整效果)</p>");
                }
                
                // 添加工作表内容
                html.AppendLine("<table>");
                
                // 计算显示的行列范围
                int startRow = 1;
                int endRow = worksheet.Dimension.End.Row;
                int startCol = 1;
                int endCol = worksheet.Dimension.End.Column;
                
                // 限制表格大小，避免生成过大的HTML
                const int maxRows = 100;
                const int maxCols = 20;
                if (endRow - startRow > maxRows) endRow = startRow + maxRows;
                if (endCol - startCol > maxCols) endCol = startCol + maxCols;
                
                // 生成表格
                for (int row = startRow; row <= endRow; row++)
                {
                    html.AppendLine("<tr>");
                    
                    for (int col = startCol; col <= endCol; col++)
                    {
                        var cell = worksheet.Cells[row, col];
                        string cellValue = cell.Text ?? "";
                        
                        // 处理合并单元格
                        string tdAttributes = "";
                        var mergeCell = worksheet.MergedCells.FirstOrDefault(m => 
                            worksheet.Cells[m].Start.Row == row && 
                            worksheet.Cells[m].Start.Column == col);
                            
                        if (!string.IsNullOrEmpty(mergeCell))
                        {
                            int rowspan = worksheet.Cells[mergeCell].End.Row - worksheet.Cells[mergeCell].Start.Row + 1;
                            int colspan = worksheet.Cells[mergeCell].End.Column - worksheet.Cells[mergeCell].Start.Column + 1;
                            
                            if (rowspan > 1) tdAttributes += $" rowspan='{rowspan}'";
                            if (colspan > 1) tdAttributes += $" colspan='{colspan}'";
                            
                            // 跳过被合并的单元格
                            if (row != worksheet.Cells[mergeCell].Start.Row || col != worksheet.Cells[mergeCell].Start.Column)
                                continue;
                        }
                        
                        // 添加样式
                        string cellStyle = "";
                        
                        // 添加字体样式
                        if (cell.Style.Font.Bold) cellStyle += "font-weight:bold;";
                        if (cell.Style.Font.Italic) cellStyle += "font-style:italic;";
                        if (cell.Style.Font.UnderLine) cellStyle += "text-decoration:underline;";
                        if (cell.Style.Font.Strike) cellStyle += "text-decoration:line-through;";
                        if (cell.Style.Font.Size > 0) cellStyle += $"font-size:{cell.Style.Font.Size}pt;";
                        
                        if (cell.Style.Font.Color.Rgb != null)
                        {
                            // 转换Excel颜色格式到CSS颜色
                            var color = cell.Style.Font.Color.Rgb;
                            // 安全检查：确保字符串长度足够
                            if (color != null && color.Length > 2)
                            {
                            cellStyle += $"color:#{color.Substring(2)};"; // 截掉前两位Alpha通道
                        }
                            else if (color != null)
                            {
                                // 如果字符串不够长，使用原始值
                                cellStyle += $"color:#{color};";
                            }
                        }
                        
                        // 添加背景颜色
                        if (cell.Style.Fill.BackgroundColor.Rgb != null)
                        {
                            var bgColor = cell.Style.Fill.BackgroundColor.Rgb;
                            // 安全检查：确保字符串长度足够
                            if (bgColor != null && bgColor.Length > 2)
                            {
                            cellStyle += $"background-color:#{bgColor.Substring(2)};";
                            }
                            else if (bgColor != null)
                            {
                                // 如果字符串不够长，使用原始值
                                cellStyle += $"background-color:#{bgColor};";
                            }
                        }
                        
                        // 对齐方式
                        switch (cell.Style.HorizontalAlignment)
                        {
                            case ExcelHorizontalAlignment.Center:
                                cellStyle += "text-align:center;";
                                break;
                            case ExcelHorizontalAlignment.Right:
                                cellStyle += "text-align:right;";
                                break;
                            case ExcelHorizontalAlignment.Left:
                                cellStyle += "text-align:left;";
                                break;
                        }
                        
                        // 垂直对齐
                        switch (cell.Style.VerticalAlignment)
                        {
                            case ExcelVerticalAlignment.Top:
                                cellStyle += "vertical-align:top;";
                                break;
                            case ExcelVerticalAlignment.Center:
                                cellStyle += "vertical-align:middle;";
                                break;
                            case ExcelVerticalAlignment.Bottom:
                                cellStyle += "vertical-align:bottom;";
                                break;
                        }
                        
                        // 应用特殊样式类
                        string cellClass = "";
                        if (cellValue.Contains("阳性")) cellClass = " class='positive'";
                        else if (cellValue.Contains("阴性")) cellClass = " class='negative'";
                        
                        // 写入单元格
                        if (!string.IsNullOrEmpty(cellStyle))
                            tdAttributes += $" style='{cellStyle}'";
                        
                        html.AppendLine($"<td{tdAttributes}{cellClass}>{cellValue}</td>");
                    }
                    
                    html.AppendLine("</tr>");
                }
                
                html.AppendLine("</table>");
                
                if (drawings.Count > 0)
                {
                    html.AppendLine("</div>");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "生成HTML表格预览时出错");
                html.AppendLine("<p style='color:red;'>生成表格预览时发生错误: " + ex.Message + "</p>");
            }
            
            html.AppendLine("</body>");
            html.AppendLine("</html>");
            
            return html.ToString();
        }
        
        /// <summary>
        /// 替换模板中的患者相关占位符
        /// </summary>
        private void ReplacePatientPlaceholders(ExcelWorksheet worksheet, string patientName, string patientCaseNumber)
        {
            _logger.LogInformation("替换患者信息占位符: 患者名={PatientName}, 病历号={CaseNumber}", 
                patientName ?? "未知患者", patientCaseNumber ?? "无病历号");

            // === 新增日志：直接检查模板中特定单元格 (例如 D6) 的内容 ===
            string targetCellAddress = "D6"; // 假设患者姓名占位符在 D6
            try
            {
                var targetCell = worksheet.Cells[targetCellAddress];
                string targetCellText = GetCellActualText(targetCell);
                _logger.LogInformation("检查模板单元格 {Address}: 预期内容='${{PatientName}}', 实际读取内容='{ActualContent}'", 
                                     targetCellAddress, targetCellText);
                
                if (targetCellText != "${PatientName}")
                {
                     _logger.LogWarning("模板单元格 {Address} 的实际内容 ('{ActualContent}') 与预期 ('${{PatientName}}') 不符！请检查模板文件。",
                                      targetCellAddress, targetCellText);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "检查模板单元格 {Address} 时发生错误。", targetCellAddress);
            }
            // === 日志结束 ===
                
            // 定义要查找和替换的占位符及其对应值
            var placeholdersToReplace = new Dictionary<string, string>
            {
                { "${PatientName}", patientName ?? "未知患者" },
                { "${PatientId}", patientCaseNumber ?? "" },
                { "${SampleId}", "" }, // 可根据需要填充默认值或从plate对象获取
                { "${CollectionDate}", DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd") }, // 可根据需要填充默认值或从plate对象获取
                { "${TestDate}", DateTime.Now.ToString("yyyy-MM-dd") }, // 可根据需要填充默认值或从plate对象获取
                { "${SampleType}", "咽拭子" }, // 添加样本类型的默认值
                { "${Gender}", "未知" }, // 添加性别默认值
                { "${Department}", "门诊" }, // 添加送检单位默认值
                { "${Section}", "呼吸科" }, // 添加送检科室默认值
                { "${Doctor}", "医生" } // 添加送检医生默认值
            };
            
            int replacedCount = 0;
            // 遍历工作表的所有单元格
            if (worksheet.Dimension != null) // 添加空值检查
            {
                 for (int i = 1; i <= worksheet.Dimension.End.Row; i++)
                 {
                     for (int j = 1; j <= worksheet.Dimension.End.Column; j++)
                     {
                         var cell = worksheet.Cells[i, j];
                         var cellText = GetCellActualText(cell); // 获取单元格文本

                         if (!string.IsNullOrEmpty(cellText))
                         {
                             string originalValue = cellText; // 保存原始值用于比较
                             string newValue = cellText;      // 初始化新值
                             bool changed = false;         // 标记是否发生替换

                             // 检查是否需要替换 ${PatientName}
                             if (placeholdersToReplace.ContainsKey("${PatientName}"))
                             {
                                 string targetPlaceholder = "${PatientName}";
                                 string replacementValue = placeholdersToReplace[targetPlaceholder];
                                 
                                 // 添加详细日志
                                 _logger.LogDebug("检查单元格 [{Row}, {Col}], 地址: {Address}, 内容: '{Content}' 是否包含 '{Placeholder}'", 
                                                  i, j, cell.Address, cellText, targetPlaceholder);

                                 if (newValue.Contains(targetPlaceholder))
                                 {
                                     newValue = newValue.Replace(targetPlaceholder, replacementValue);
                                     changed = true;
                                     _logger.LogInformation("成功匹配并准备替换单元格 [{Row}, {Col}] 中的 '{Placeholder}' 为 '{Replacement}'", 
                                                          i, j, targetPlaceholder, replacementValue);
                                 }
                             }
                             
                             // 遍历其他占位符并进行替换
                             foreach(var placeholderPair in placeholdersToReplace)
                             {
                                 // 跳过已处理的PatientName
                                 if (placeholderPair.Key == "${PatientName}") continue;

                                 if (newValue.Contains(placeholderPair.Key))
                                 {
                                     newValue = newValue.Replace(placeholderPair.Key, placeholderPair.Value);
                                     changed = true;
                                     _logger.LogDebug("替换单元格 [{Row}, {Col}] 中的 '{OldPlaceholder}' 为 '{NewValue}'", 
                                                      i, j, placeholderPair.Key, placeholderPair.Value);
                                 }
                             }

                             // 如果发生了替换，则更新单元格值并计数
                             if (changed)
                             {
                                 cell.Value = newValue;
                                 replacedCount++;
                             }
                         }
                     }
                 }
            }
            else
            {
                _logger.LogWarning("Worksheet dimension is null, cannot replace placeholders.");
            }
            _logger.LogInformation("完成患者信息占位符替换，共替换 {Count} 处", replacedCount);
        }

        /// <summary>
        /// 替换模板中的通用占位符
        /// </summary>
        private async Task ReplaceTemplatePlaceholders(ExcelWorksheet worksheet, Plate plate)
        {
            _logger.LogInformation("替换模板中的通用占位符");
            
            // 简化占位符映射
            var placeholderPairs = new Dictionary<string, string>
            {
                // 板相关信息
                { "[[PlateID]]", plate.Id.ToString() },
                { "[[PlateName]]", plate.Name },
                { "[板号]", plate.Id.ToString() },
                { "[板名称]", plate.Name },
                { "${PlateId}", plate.Id.ToString() },
                { "${PlateName}", plate.Name },
                
                // 日期信息
                { "[[Date]]", DateTime.Now.ToString("yyyy-MM-dd") },
                { "${ReportDate}", DateTime.Now.ToString("yyyy-MM-dd") },
                { "${ReportTime}", DateTime.Now.ToString("HH:mm:ss") },
                { "[报告日期]", DateTime.Now.ToString("yyyy-MM-dd") },
                { "[报告时间]", DateTime.Now.ToString("HH:mm:ss") },
                
                // 统计信息
                { "[[TotalSamples]]", plate.WellLayouts.Count.ToString() },
                { "[[PositiveCount]]", plate.WellLayouts.Count(w => IsPositiveResult(w)).ToString() },
                { "[[NegativeCount]]", plate.WellLayouts.Count(w => IsNegativeResult(w)).ToString() },
                { "${TotalSamples}", plate.WellLayouts.Count.ToString() },
                { "${PositiveCount}", plate.WellLayouts.Count(w => IsPositiveResult(w)).ToString() },
                { "${NegativeCount}", plate.WellLayouts.Count(w => IsNegativeResult(w)).ToString() },
                
                // 患者信息默认值（如果没有通过ReplacePatientPlaceholders设置）
                { "${PatientName}", "未知患者" },
                { "${PatientId}", "" },
                { "${Gender}", "未知" },
                { "${SampleId}", "" },
                { "${SampleType}", "咽拭子" },
                { "${CollectionDate}", DateTime.Now.AddDays(-1).ToString("yyyy-MM-dd") },
                { "${TestDate}", DateTime.Now.ToString("yyyy-MM-dd") },
                { "${Department}", "门诊" },
                { "${Section}", "呼吸科" },
                { "${Doctor}", "医生" },
                
                // 数据行占位符（将在填充实际数据时替换）
                { "${Target}", "" },
                { "${Concentration}", "" },
                { "${CtValue}", "" },
                { "${Result}", "" }
            };
            
            // 遍历所有单元格，直接替换已知的占位符
            for (int i = 1; i <= worksheet.Dimension.End.Row; i++)
            {
                for (int j = 1; j <= worksheet.Dimension.End.Column; j++)
                {
                    var cell = worksheet.Cells[i, j];
                    var cellText = GetCellActualText(cell);
                    if (string.IsNullOrEmpty(cellText)) continue;
                    
                    // 跳过包含[[DataStart]]标记的单元格，这些将在数据填充阶段处理
                    if (cellText.Contains("[[DataStart]]")) continue;
                    
                    string newValue = cellText;
                    bool changed = false;
                    
                    // 替换所有已知占位符
                    foreach (var pair in placeholderPairs)
                    {
                        if (newValue.Contains(pair.Key))
                        {
                            newValue = newValue.Replace(pair.Key, pair.Value);
                            changed = true;
                        }
                    }
                    
                    // 更新单元格
                    if (changed)
                    {
                        cell.Value = newValue;
                        _logger.LogDebug("替换了单元格 [{Row}, {Col}] 的占位符", i, j);
                    }
                }
            }
        }

        /// <summary>
        /// 填充整板报告数据
        /// </summary>
        private async Task FillPlateReportData(ExcelWorksheet worksheet, Plate plate)
        {
            _logger.LogInformation("正在填充整板报告数据");
            
            // 1. 先替换基本占位符
            for (int i = 1; i <= worksheet.Dimension.End.Row; i++)
            {
                for (int j = 1; j <= worksheet.Dimension.End.Column; j++)
                {
                    var cell = worksheet.Cells[i, j];
                    string cellText = GetCellActualText(cell);
                    
                    if (string.IsNullOrEmpty(cellText)) continue;
                    
                    // 检查是否为DataStart标记 - 先处理这种情况
                    if (cellText == "[[DataStart]]" || cellText.Trim() == "[[DataStart]]" || 
                        cellText == "[DataStart]" || cellText.Contains("DataStart"))
                    {
                        _logger.LogInformation("在位置 [{Row}, {Col}] 找到数据起始标记，将开始填充数据", i, j);
                        cell.Value = null; // 清除标记
                        continue;
                    }
                    
                    // 尝试替换基本占位符，板名称、板ID等
                    string newValue = cellText;
                    
                    // 替换报告通用信息
                    if (newValue.Contains("${PlateId}") || newValue.Contains("[[PlateID]]") || newValue.Contains("[板号]"))
                        newValue = newValue.Replace("${PlateId}", plate.Id.ToString())
                                        .Replace("[[PlateID]]", plate.Id.ToString())
                                        .Replace("[板号]", plate.Id.ToString());
                    
                    if (newValue.Contains("${PlateName}") || newValue.Contains("[[PlateName]]") || newValue.Contains("[板名称]"))
                        newValue = newValue.Replace("${PlateName}", plate.Name)
                                        .Replace("[[PlateName]]", plate.Name)
                                        .Replace("[板名称]", plate.Name);
                    
                    if (newValue.Contains("${ReportDate}") || newValue.Contains("[报告日期]"))
                        newValue = newValue.Replace("${ReportDate}", DateTime.Now.ToString("yyyy-MM-dd"))
                                        .Replace("[报告日期]", DateTime.Now.ToString("yyyy-MM-dd"));
                    
                    // 更新单元格
                    if (newValue != cellText)
                    {
                        cell.Value = newValue;
                    }
                }
            }
            
            // 2. 寻找数据表格起始位置
            int dataStartRow = -1;
            int dataStartCol = -1;
            
            // 找到 [[DataStart]] 标记 - 尝试多种格式
            for (int i = 1; i <= worksheet.Dimension.End.Row; i++)
            {
                for (int j = 1; j <= worksheet.Dimension.End.Column; j++)
                {
                    var cellText = GetCellActualText(worksheet.Cells[i, j]);
                    // 查找各种可能的DataStart标记
                    if (!string.IsNullOrEmpty(cellText) && 
                        (cellText.Equals("[[DataStart]]") || 
                         cellText.Trim().Equals("[[DataStart]]") || 
                         cellText.Equals("[DataStart]") ||
                         cellText.Contains("DataStart")))
                    {
                        dataStartRow = i;
                        dataStartCol = j;
                        // 清除标记
                        worksheet.Cells[i, j].Value = null;
                        _logger.LogInformation("找到整板报告数据起始标记，位置: [{Row}, {Col}]", i, j);
                        break;
                    }
                }
                if (dataStartRow != -1) break;
            }
            
            // 如果没有标记，尝试查找表头行
            if (dataStartRow == -1)
            {
                for (int i = 1; i <= worksheet.Dimension.End.Row; i++)
                {
                    bool foundHeader = false;
                    for (int j = 1; j <= worksheet.Dimension.End.Column; j++)
                    {
                        var cellText = worksheet.Cells[i, j].Text;
                        if (!string.IsNullOrEmpty(cellText) && 
                            (cellText.Contains("孔位") || 
                             cellText.Contains("Well") || 
                             cellText.Contains("通道") || 
                             cellText.Contains("CT值") ||
                             cellText.Contains("靶标") ||
                             cellText.Contains("结果") ||
                             cellText.Contains("阳性") ||
                             cellText.Contains("阴性")))
                        {
                            foundHeader = true;
                            dataStartRow = i + 1; // 数据从表头的下一行开始
                            dataStartCol = 1;
                            _logger.LogInformation("通过表头找到数据位置，开始行: {StartRow}", dataStartRow);
                            break;
                        }
                    }
                    if (foundHeader) break;
                }
            }
            
            // 3. 如果找到了数据位置，开始填充数据
            if (dataStartRow > 0)
            {
                _logger.LogInformation("开始在行 {Row} 填充整板数据", dataStartRow);
                
                // 查找表格列的位置
                int wellCol = -1;      // 孔位列
                int patientCol = -1;   // 患者姓名列
                int caseNumCol = -1;   // 病例号列
                int channelCol = -1;   // 通道列
                int targetCol = -1;    // 靶标列
                int ctCol = -1;        // CT值列
                int resultCol = -1;    // 结果列
                int concCol = -1;      // 浓度列
                int yangxingCol = -1;  // 阳性列
                
                // 在表头行查找各列
                int headerRow = Math.Max(1, dataStartRow - 1);
                for (int j = 1; j <= worksheet.Dimension.End.Column; j++)
                {
                    string headerText = GetCellActualText(worksheet.Cells[headerRow, j]);
                    if (string.IsNullOrEmpty(headerText)) continue;
                    
                    _logger.LogDebug("表头列 {Col}: '{Text}'", j, headerText);
                    
                    if (headerText.Contains("孔位") || headerText.Contains("Well"))
                        wellCol = j;
                    else if (headerText.Contains("患者") || headerText.Contains("姓名"))
                        patientCol = j;
                    else if (headerText.Contains("病历") || headerText.Contains("ID"))
                        caseNumCol = j;
                    else if (headerText.Contains("通道") || headerText.Contains("Channel"))
                        channelCol = j;
                    else if (headerText.Contains("靶标") || headerText.Contains("病原体") || headerText.Equals("CY5"))
                        targetCol = j;
                    else if (headerText.Contains("CT") || headerText.Contains("Ct"))
                        ctCol = j;
                    else if (headerText.Contains("结果") || headerText.Contains("检测结果"))
                        resultCol = j;
                    else if (headerText.Contains("浓度"))
                        concCol = j;
                    else if (headerText.Equals("阳性") || headerText.Contains("阳性"))
                        yangxingCol = j;
                }
                
                _logger.LogInformation("找到列: 孔位={WellCol}, 患者={PatientCol}, 病历号={CaseNumberCol}, 通道={ChannelCol}, 靶标={TargetCol}, CT值={CtValueCol}, 结果={ResultCol}", 
                    wellCol, patientCol, caseNumCol, channelCol, targetCol, ctCol, resultCol);
                
                // 设置默认列位置（如果无法找到）
                if (wellCol < 0) wellCol = 1;
                if (targetCol < 0 && channelCol > 0) targetCol = channelCol + 1;
                if (channelCol < 0 && targetCol > 0) channelCol = targetCol - 1;
                if (ctCol < 0 && resultCol > 0) ctCol = resultCol - 1;
                if (resultCol < 0 && ctCol > 0) resultCol = ctCol + 1;
                
                // 不再按照靶标分组，而是保留原始孔位中的所有荧光通道
                var wellsData = plate.WellLayouts
                    .OrderBy(w => w.Row)
                    .ThenBy(w => w.Column)
                    .ToList();
                
                int currentRow = dataStartRow;
                
                _logger.LogInformation("准备填充 {Count} 个孔位数据", wellsData.Count);
                
                // 如果没有数据，添加示例数据
                if (wellsData.Count == 0)
                {
                    _logger.LogWarning("未找到孔位数据，将添加示例数据");
                    
                    if (wellCol > 0)
                        worksheet.Cells[currentRow, wellCol].Value = "A1";
                    
                    if (patientCol > 0)
                        worksheet.Cells[currentRow, patientCol].Value = "未知患者";
                        
                    if (targetCol > 0)
                        worksheet.Cells[currentRow, targetCol].Value = "CY5";
                    
                    if (ctCol > 0)
                        worksheet.Cells[currentRow, ctCol].Value = 0.0;
                    
                    if (resultCol > 0)
                        worksheet.Cells[currentRow, resultCol].Value = "阴性";
                    
                    currentRow++;
                }
                
                // 不再分组数据，保留所有荧光通道
                foreach (var well in wellsData)
                {
                    // 跳过没有有效通道或靶标的孔位
                    if (string.IsNullOrEmpty(well.Channel) && string.IsNullOrEmpty(well.TargetName)) 
                        continue;
                    
                    // 填充基础数据
                    if (wellCol > 0)
                        worksheet.Cells[currentRow, wellCol].Value = well.WellName;
                    
                    if (patientCol > 0)
                        worksheet.Cells[currentRow, patientCol].Value = string.IsNullOrEmpty(well.PatientName) ? "未知患者" : well.PatientName;
                    
                    if (caseNumCol > 0)
                        worksheet.Cells[currentRow, caseNumCol].Value = well.PatientCaseNumber;
                    
                    if (channelCol > 0)
                        worksheet.Cells[currentRow, channelCol].Value = well.Channel;
                    
                    if (targetCol > 0)
                        worksheet.Cells[currentRow, targetCol].Value = well.TargetName;
                    
                    if (ctCol > 0)
                    {
                        // 保持原始精度，不截断CT值
                        if (well.CtValue.HasValue)
                        {
                            // 使用指定的数字格式确保保留小数点后两位
                            worksheet.Cells[currentRow, ctCol].Value = well.CtValue.Value;
                            worksheet.Cells[currentRow, ctCol].Style.Numberformat.Format = "0.00";
                        }
                        else
                        {
                            worksheet.Cells[currentRow, ctCol].Value = "";
                        }
                    }
                    
                    if (resultCol > 0)
                    {
                        // 获取结果文本
                        string resultText = GetResultText(well);
                        worksheet.Cells[currentRow, resultCol].Value = resultText;
                        
                        // 如果是阳性结果，设置字体颜色为红色
                        if (IsPositiveResult(well))
                        {
                            worksheet.Cells[currentRow, resultCol].Style.Font.Color.SetColor(Color.Red);
                            worksheet.Cells[currentRow, resultCol].Style.Font.Bold = true;
                        }
                    }
                    
                    // 填充浓度值或阳性状态
                    if (concCol > 0)
                    {
                        // 使用分析方法配置中的公式计算浓度
                        string concValue = CalculateConcentration(plate, well.CtValue, well.Channel, well.TargetName);
                        worksheet.Cells[currentRow, concCol].Value = concValue;
                    }
                    
                    if (yangxingCol > 0)
                    {
                        // 填充阳性状态
                        string yangxingStatus = IsPositiveResult(well) ? "阳性" : "阴性";
                        worksheet.Cells[currentRow, yangxingCol].Value = yangxingStatus;
                        
                        // 设置阳性为红色
                        if (IsPositiveResult(well))
                        {
                            worksheet.Cells[currentRow, yangxingCol].Style.Font.Color.SetColor(Color.Red);
                            worksheet.Cells[currentRow, yangxingCol].Style.Font.Bold = true;
                        }
                    }
                    
                    // 应用样式 - 添加边框
                    for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                    {
                        worksheet.Cells[currentRow, col].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells[currentRow, col].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells[currentRow, col].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        worksheet.Cells[currentRow, col].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    }
                    
                    currentRow++;
                }
                
                // 设置表格样式
                if (currentRow > dataStartRow)
                {
                    var tableRange = worksheet.Cells[headerRow, 1, currentRow - 1, worksheet.Dimension.End.Column];
                    
                    // 设置表格边框
                    tableRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    tableRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    tableRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    tableRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    
                    // 设置表头样式
                    var headerRange = worksheet.Cells[headerRow, 1, headerRow, worksheet.Dimension.End.Column];
                    headerRange.Style.Font.Bold = true;
                    headerRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    headerRange.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                }
                
                _logger.LogInformation("填充完成，共填充 {Count} 行数据", currentRow - dataStartRow);
            }
            else
            {
                _logger.LogWarning("无法找到数据表格位置，无法填充整板报告数据");
            }
            
            await Task.CompletedTask;
        }

        /// <summary>
        /// 填充患者报告数据
        /// </summary>
        private async Task FillPatientReportData(ExcelWorksheet worksheet, Plate plate)
        {
            _logger.LogInformation("开始填充患者报告数据");
            // 实现患者报告数据填充逻辑
            await Task.CompletedTask;
        }
        
        /// <summary>
        /// 检查结果是否为阳性
        /// </summary>
        private bool IsPositiveResult(WellLayout well)
        {
            // 如果有CT值，且在合理范围内（大于0且小于38），则认为是阳性
            if (well.CtValue.HasValue && well.CtValue.Value > 0 && well.CtValue.Value < 38)
            {
                return true;
            }
            
            // 判断特殊标记，如果没有特殊标记或特殊标记不是"阴性"相关的标记，可能是阳性
            if (!string.IsNullOrEmpty(well.CtValueSpecialMark) && 
                (well.CtValueSpecialMark.Contains("阳性") || well.CtValueSpecialMark.Equals("+", StringComparison.OrdinalIgnoreCase)))
            {
                return true;
            }
            
            // 如果有标准浓度且大于0，可能是阳性
            if (well.StandardConcentration.HasValue && well.StandardConcentration.Value > 0)
            {
                return true;
            }
            
            return false;
        }
        
        /// <summary>
        /// 检查结果是否为阴性
        /// </summary>
        private bool IsNegativeResult(WellLayout well)
        {
            return !IsPositiveResult(well);
        }
        
        /// <summary>
        /// 获取结果文本描述
        /// </summary>
        private string GetResultText(WellLayout well)
        {
            if (IsPositiveResult(well))
            {
                return "阳性";
            }
            else
            {
                return "阴性";
            }
        }
        
        /// <summary>
        /// 获取Excel单元格的实际文本内容，兼容各种可能的内容格式
        /// </summary>
        private string GetCellActualText(ExcelRangeBase cell)
        {
            if (cell == null) return string.Empty;
            
            // 尝试多种方式获取单元格内容
            string text = string.Empty;
            
            // 首先尝试直接获取Text属性
            text = cell.Text;
            if (!string.IsNullOrEmpty(text))
                return text;
            
            // 如果Text为空，尝试获取Value属性
            if (cell.Value != null)
            {
                text = cell.Value.ToString();
                if (!string.IsNullOrEmpty(text))
                    return text;
            }
            
            // 如果仍然为空，尝试通过公式获取
            if (!string.IsNullOrEmpty(cell.Formula))
            {
                text = cell.Formula;
                _logger.LogDebug("通过公式获取文本: {Formula}", text);
                return text;
            }
            
            return string.Empty;
        }

        private async Task AutoFixTemplateIfNeededAsync(ReportTemplate template)
        {
            try
            {
                _logger.LogInformation("尝试自动修复模板: {TemplatePath}", template.FilePath);
                
                // 设置EPPlus许可证上下文
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                
                // 创建备份
                string backupPath = $"{template.FilePath}.autofix.backup";
                if (!File.Exists(backupPath))
                {
                    File.Copy(template.FilePath, backupPath, false);
                    _logger.LogInformation("已创建模板备份: {BackupPath}", backupPath);
                }
                
                bool needsSave = false;
                
                // 打开Excel文件
                using (var package = new ExcelPackage(new FileInfo(template.FilePath)))
                {
                    foreach (var worksheet in package.Workbook.Worksheets)
                    {
                        if (worksheet.Dimension == null) continue;
                        
                        bool hasDataStart = false;
                        bool foundTableStructure = false;
                        
                        // 1. 检查是否有[[DataStart]]标记
                        for (int i = 1; i <= worksheet.Dimension.End.Row; i++)
                        {
                            for (int j = 1; j <= worksheet.Dimension.End.Column; j++)
                            {
                                var cellValue = GetCellActualText(worksheet.Cells[i, j]);
                                if (!string.IsNullOrEmpty(cellValue) && cellValue.Contains("[[DataStart]]"))
                                {
                                    hasDataStart = true;
                                    break;
                                }
                            }
                            if (hasDataStart) break;
                        }
                        
                        // 2. 如果没有[[DataStart]]标记，尝试找到表格结构并添加标记
                        if (!hasDataStart)
                        {
                            _logger.LogInformation("模板中未找到[[DataStart]]标记，尝试自动添加");
                            
                            // 查找可能的表头
                            for (int row = 1; row <= worksheet.Dimension.End.Row; row++)
                            {
                                bool potentialHeader = false;
                                int headerCol = 1;
                                
                                for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                                {
                                    string cellText = worksheet.Cells[row, col].Text;
                                    if (!string.IsNullOrEmpty(cellText) && 
                                        (cellText.Contains("病原体") || 
                                         cellText.Contains("CT值") || 
                                         cellText.Contains("靶标") ||
                                         cellText.Contains("结果")))
                                    {
                                        potentialHeader = true;
                                        headerCol = col;
                                        break;
                                    }
                                }
                                
                                if (potentialHeader && row < worksheet.Dimension.End.Row)
                                {
                                    // 在表头下方添加[[DataStart]]标记
                                    worksheet.Cells[row + 1, headerCol].Value = "[[DataStart]]";
                                    _logger.LogInformation("在位置 [{Row}, {Col}] 自动添加了[[DataStart]]标记", row + 1, headerCol);
                                    needsSave = true;
                                    foundTableStructure = true;
                                    break;
                                }
                            }
                            
                            // 如果找不到表头，尝试在第15行添加
                            if (!foundTableStructure)
                            {
                                int targetRow = Math.Min(15, worksheet.Dimension.End.Row);
                                worksheet.Cells[targetRow, 1].Value = "[[DataStart]]";
                                _logger.LogInformation("未找到表结构，在位置 [{Row}, 1] 强制添加了[[DataStart]]标记", targetRow);
                                needsSave = true;
                            }
                        }
                        
                        // 3. 替换旧格式占位符
                        string[,] placeholders = new string[,]
                        {
                            { "[受检人姓名]", "${PatientName}" },
                            { "[受检人ID]", "${PatientId}" },
                            { "[受检人性别]", "${Gender}" },
                            { "[病历号]", "${PatientId}" },
                            { "[样本条码]", "${SampleId}" },
                            { "[样本编码]", "${SampleCode}" },
                            { "[接收日期]", "${CollectionDate}" },
                            { "[采样日期]", "${TestDate}" },
                            { "[样本类型]", "${SampleType}" },
                            { "[送检单位]", "${Department}" },
                            { "[送检科室]", "${Section}" },
                            { "[送检医生]", "${Doctor}" },
                            { "[病原体名]", "${Target}" },
                            { "[浓度值]", "${Concentration}" },
                            { "[CT值]", "${CtValue}" },
                            { "[检测结果]", "${Result}" }
                        };
                        
                        for (int i = 1; i <= worksheet.Dimension.End.Row; i++)
                        {
                            for (int j = 1; j <= worksheet.Dimension.End.Column; j++)
                            {
                                var cellValue = worksheet.Cells[i, j].Text;
                                if (!string.IsNullOrEmpty(cellValue))
                                {
                                    string newValue = cellValue;
                                    for (int k = 0; k < placeholders.GetLength(0); k++)
                                    {
                                        if (newValue.Contains(placeholders[k, 0]))
                                        {
                                            newValue = newValue.Replace(placeholders[k, 0], placeholders[k, 1]);
                                            needsSave = true;
                                        }
                                    }
                                    
                                    if (newValue != cellValue)
                                    {
                                        worksheet.Cells[i, j].Value = newValue;
                                        _logger.LogDebug("替换了单元格 [{Row}, {Col}] 中的占位符", i, j);
                                    }
                                }
                            }
                        }
                    }
                    
                    // 如果有修改，保存文件
                    if (needsSave)
                    {
                        _logger.LogInformation("模板已修改，正在保存修复后的模板");
                        await package.SaveAsync();
                    }
                    else
                    {
                        _logger.LogInformation("模板检查完成，无需修复");
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "自动修复模板时出错: {Message}", ex.Message);
                // 出错时不抛出异常，继续使用原始模板
            }
        }

        public Task<IEnumerable<ReportTemplate>> GetReportTemplatesAsync()
        {
            _logger.LogInformation("获取报告模板列表");
            // 扫描模板目录
            string templatesDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates");
            var templates = new List<ReportTemplate>();
            
            if (Directory.Exists(templatesDir))
            {
                var templateFiles = Directory.GetFiles(templatesDir, "*.xlsx");
                foreach (var file in templateFiles)
                {
                    // 跳过以~$开头的临时文件（Excel打开文件时创建的临时文件）
                    if (Path.GetFileName(file).StartsWith("~$")) continue;
                    
                    templates.Add(new ReportTemplate 
                    { 
                        Id = Guid.NewGuid(), 
                        Name = Path.GetFileNameWithoutExtension(file), 
                        FilePath = file, 
                        IsExcelTemplate = true,
                        Description = "从模板目录自动加载" // 可以在这里添加更多描述信息
                    });
                }
            }
            else
            {
                _logger.LogWarning("模板目录不存在: {TemplatesDir}", templatesDir);
                // 创建模板目录
                try
                {
                    Directory.CreateDirectory(templatesDir);
                    _logger.LogInformation("创建了模板目录: {TemplatesDir}", templatesDir);
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "创建模板目录失败: {TemplatesDir}", templatesDir);
                }
            }
            
            return Task.FromResult<IEnumerable<ReportTemplate>>(templates);
        }

        public async Task<ReportTemplate?> GetReportTemplateAsync(Guid id)
        {
             _logger.LogInformation("获取指定ID的报告模板: ID={TemplateId}", id);
             var templates = await GetReportTemplatesAsync();
            return templates.FirstOrDefault(t => t.Id == id);
        }

        public Task<ReportTemplate> SaveReportTemplateAsync(ReportTemplate reportTemplate)
        {
             _logger.LogInformation("保存报告模板: Name={TemplateName}", reportTemplate.Name);
            // 目前仅返回原模板，后续可以实现保存逻辑
            return Task.FromResult(reportTemplate);
        }

        public Task<bool> ValidateReportTemplateAsync(string templateFilePath)
        {
            _logger.LogInformation("校验报告模板: {TemplatePath}", templateFilePath);
            bool isValid = false;
            if (!File.Exists(templateFilePath))
            {
                 _logger.LogWarning("模板文件不存在");
                 return Task.FromResult(false);
            }
            try
            {
                // 使用EPPlus尝试打开文件进行基本验证
                using (var package = new ExcelPackage(new FileInfo(templateFilePath)))
                {
                    isValid = package.Workbook.Worksheets.Any(); // 检查是否至少有一个工作表
                }
                 _logger.LogInformation("模板校验结果: {IsValid}", isValid);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "校验模板时出错");
                isValid = false;
            }
            return Task.FromResult(isValid);
        }

        /// <summary>
        /// 查找并处理患者报告中的病原体列表表格（使用分析结果）
        /// </summary>
        private async Task FindAndProcessPatientResults(ExcelWorksheet worksheet, IEnumerable<AnalysisResultItem> analysisResults)
        {
            _logger.LogInformation("开始查找和处理病原体列表表格 (使用分析结果)");

            // --- 查找表头行 --- 
            // 修改：不再查找合并的标题行，而是查找包含特定列名的实际表头行
            int tableHeaderRow = -1;
            string[] expectedHeaders = { "病原体", "浓度", "CT", "结果" }; // 简化的关键词
            int headerMatchCountThreshold = 2; // 至少需要匹配到2个关键词才认为是表头

            for (int i = 1; i <= worksheet.Dimension.End.Row; i++)
            {
                int matchCount = 0;
                for (int j = 1; j <= worksheet.Dimension.End.Column; j++)
                {
                    var cellText = GetCellActualText(worksheet.Cells[i, j]);
                    if (!string.IsNullOrEmpty(cellText))
                    {
                        // 检查单元格文本是否包含任何预期的表头关键词
                        if (expectedHeaders.Any(h => cellText.Contains(h)))
                        {
                            matchCount++;
                        }
                    }
                }
                // 如果当前行匹配到了足够数量的关键词，则认为是表头行
                if (matchCount >= headerMatchCountThreshold)
                {
                    tableHeaderRow = i;
                    _logger.LogInformation("通过关键词匹配找到可能的表头行：{Row}", tableHeaderRow);
                    break;
                }
            }

            if (tableHeaderRow > 0)
            {
                _logger.LogInformation("开始处理病原体列表，表头行：{Row}", tableHeaderRow);

                // --- 查找列位置 (在找到的表头行 tableHeaderRow 中查找) ---
                int pathogenCol = -1, concentrationCol = -1, ctValueCol = -1, resultCol = -1;
                for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                {
                    var headerText = GetCellActualText(worksheet.Cells[tableHeaderRow, col]);
                    if (string.IsNullOrEmpty(headerText)) continue;

                    // 使用更精确的匹配逻辑
                    if ((headerText.Equals("病原体名") || headerText.Contains("病原体") || headerText.Contains("靶标")) && pathogenCol == -1)
                        pathogenCol = col;
                    else if ((headerText.Contains("浓度") || headerText.Contains("拷贝/mL")) && concentrationCol == -1)
                        concentrationCol = col;
                    else if ((headerText.Equals("CT值") || (headerText.Contains("CT") && !headerText.Contains("Concentration"))) && ctValueCol == -1) // 避免匹配浓度 Concentration
                        ctValueCol = col;
                    else if ((headerText.Contains("检测结果") || headerText.Contains("结果") || headerText.Contains("结论")) && resultCol == -1)
                        resultCol = col;
                }
                _logger.LogInformation("表头行查找结果：病原体={0}, 浓度={1}, CT值={2}, 结果={3}", pathogenCol, concentrationCol, ctValueCol, resultCol);

                // --- 查找数据起始行 (从表头下一行开始查找 [[DataStart]]) ---
                int dataStartRow = -1, dataStartCol = -1;
                for (int i = tableHeaderRow + 1; i <= Math.Min(tableHeaderRow + 5, worksheet.Dimension.End.Row); i++)
                { 
                    for (int j = 1; j <= worksheet.Dimension.End.Column; j++)
                    {
                        var cellText = GetCellActualText(worksheet.Cells[i, j]);
                        if (!string.IsNullOrEmpty(cellText) && cellText.Contains("[[DataStart]]"))
                        {
                            dataStartRow = i; dataStartCol = j;
                            _logger.LogInformation("找到[[DataStart]]标记：位置 [{Row}, {Col}]", dataStartRow, dataStartCol);
                            // 在DataStart行再次尝试通过占位符确认/查找列位置 (作为后备)
                            if (pathogenCol < 1) pathogenCol = FindColumnByPlaceholder(worksheet, dataStartRow, "${Target}") ?? pathogenCol;
                            if (concentrationCol < 1) concentrationCol = FindColumnByPlaceholder(worksheet, dataStartRow, "${Concentration}") ?? concentrationCol;
                            if (ctValueCol < 1) ctValueCol = FindColumnByPlaceholder(worksheet, dataStartRow, "${CtValue}") ?? ctValueCol;
                            if (resultCol < 1) resultCol = FindColumnByPlaceholder(worksheet, dataStartRow, "${Result}") ?? resultCol;
                            break;
                        }
                    } 
                    if (dataStartRow != -1) break; 
                }
                // 如果未找到DataStart标记，默认数据从表头下一行开始
                if (dataStartRow == -1) { dataStartRow = tableHeaderRow + 1; _logger.LogWarning("未找到[[DataStart]]标记，使用表头下一行 {Row} 作为数据起始行", dataStartRow); }
                
                 // --- 再次确认列位置 (如果通过表头未找到，强制使用占位符位置) ---
                 if (pathogenCol < 1) pathogenCol = FindColumnByPlaceholder(worksheet, dataStartRow, "${Target}") ?? 1;
                 if (concentrationCol < 1) concentrationCol = FindColumnByPlaceholder(worksheet, dataStartRow, "${Concentration}") ?? 2;
                 if (ctValueCol < 1) ctValueCol = FindColumnByPlaceholder(worksheet, dataStartRow, "${CtValue}") ?? 3;
                 if (resultCol < 1) resultCol = FindColumnByPlaceholder(worksheet, dataStartRow, "${Result}") ?? 4;
                 _logger.LogInformation("最终列位置：病原体={0}, 浓度={1}, CT值={2}, 结果={3}", pathogenCol, concentrationCol, ctValueCol, resultCol);


                // --- 保存样式、删除示例行、处理模板行 --- (代码与上次修改相同，此处省略详细代码)
                var templateRowStyles = new Dictionary<int, ExcelStyle>();
                for (int col = 1; col <= worksheet.Dimension.End.Column; col++) { /* 保存样式 */ templateRowStyles[col] = worksheet.Cells[dataStartRow, col].Style; }
                int lastRowToDelete = dataStartRow;
                for (int i = dataStartRow + 1; i <= worksheet.Dimension.End.Row; i++) { /* 查找结束位置 */ bool isEmpty = true; for (int j = 1; j <= worksheet.Dimension.End.Column; j++) { if (!string.IsNullOrEmpty(GetCellActualText(worksheet.Cells[i,j]))) {isEmpty = false; break;} } if (isEmpty) break; lastRowToDelete = i; }
                int rowsToDelete = lastRowToDelete - dataStartRow;
                if (rowsToDelete > 0) { /* 删除行 */ worksheet.DeleteRow(dataStartRow + 1, rowsToDelete); _logger.LogInformation("删除 {Count} 行示例数据", rowsToDelete); }
                bool templateRowProcessed = false;
                if (dataStartCol > 0) { /* 处理 [[DataStart]] */ var cell = worksheet.Cells[dataStartRow, dataStartCol]; string cellText = GetCellActualText(cell); if (!string.IsNullOrEmpty(cellText) && cellText.Contains("[[DataStart]]")) { cell.Value = cellText.Replace("[[DataStart]]", "").Trim(); templateRowProcessed = true; } }
                string[] placeholders = { "${Target}", "${Concentration}", "${CtValue}", "${Result}" };
                for(int col = 1; col <= worksheet.Dimension.End.Column; col++) { /* 清除占位符 */ var cell = worksheet.Cells[dataStartRow, col]; var cellText = GetCellActualText(cell); if (!string.IsNullOrEmpty(cellText) && placeholders.Any(p => cellText.Contains(p))) { cell.Value = null; _logger.LogDebug("清除了模板行 [{Row}, {Col}] 的占位符", dataStartRow, col); templateRowProcessed = true; } }

                // --- 准备填充数据 ---
                var positiveResults = analysisResults
                    .Where(r => r.DetectionResult?.Equals("阳性", StringComparison.OrdinalIgnoreCase) == true)
                    .OrderBy(r => r.TargetName) 
                    .ToList();
                _logger.LogInformation("筛选出 {Count} 个阳性检测结果用于填充表格", positiveResults.Count);

                // --- 处理无阳性结果的情况 ---
                if (positiveResults.Count == 0) { /* 添加未检出/阴性行 */ _logger.LogWarning("无阳性结果，添加默认'未检出'行"); if (pathogenCol > 0) worksheet.Cells[dataStartRow, pathogenCol].Value = "未检出"; if (ctValueCol > 0) worksheet.Cells[dataStartRow, ctValueCol].Value = "-"; if (concentrationCol > 0) worksheet.Cells[dataStartRow, concentrationCol].Value = "-"; if (resultCol > 0) { worksheet.Cells[dataStartRow, resultCol].Value = "阴性"; /* ...设置样式...*/ } }
                 else {
                    // --- 填充阳性结果数据 ---
                    bool isFirstRow = true;
                    int currentRow = dataStartRow;
                    foreach (var resultItem in positiveResults) {
                        if (!isFirstRow) { 
                             // 插入新行
                             worksheet.InsertRow(currentRow + 1, 1);
                             currentRow++;
                             // 复制样式 (更全面的复制)
                             for (int col = 1; col <= worksheet.Dimension.End.Column; col++) {
                                 if (templateRowStyles.ContainsKey(col)) {
                                     var targetCell = worksheet.Cells[currentRow, col];
                                     var sourceStyle = templateRowStyles[col];
                                     
                                     // 尝试复制更多样式属性
                                     targetCell.Style.Font.Name = sourceStyle.Font.Name;
                                     targetCell.Style.Font.Size = sourceStyle.Font.Size;
                                     targetCell.Style.Font.Bold = sourceStyle.Font.Bold;
                                     targetCell.Style.Font.Italic = sourceStyle.Font.Italic;
                                     targetCell.Style.Font.UnderLine = sourceStyle.Font.UnderLine;
                                     // 修改：解析颜色字符串
                                     if (!string.IsNullOrEmpty(sourceStyle.Font.Color.Rgb))
                                     {
                                         try {
                                            targetCell.Style.Font.Color.SetColor(System.Drawing.ColorTranslator.FromHtml("#" + sourceStyle.Font.Color.Rgb.Substring(2))); // 尝试从 RGB 部分创建颜色
                                         }
                                         catch (Exception ex) {
                                             _logger.LogWarning(ex, "无法解析字体颜色 RGB: {RgbColor}", sourceStyle.Font.Color.Rgb);
                                             // 可以设置默认颜色或跳过
                                         }
                                     }
                                     // targetCell.Style.Font.Color.SetColor(sourceStyle.Font.Color.Rgb); // 旧的错误代码
                                     
                                     targetCell.Style.Fill.PatternType = sourceStyle.Fill.PatternType;
                                     if(sourceStyle.Fill.PatternType != ExcelFillStyle.None && !string.IsNullOrEmpty(sourceStyle.Fill.BackgroundColor.Rgb))
                                     { 
                                        // 修改：解析颜色字符串
                                        try {
                                            targetCell.Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#" + sourceStyle.Fill.BackgroundColor.Rgb.Substring(2)));
                                        }
                                        catch (Exception ex) {
                                             _logger.LogWarning(ex, "无法解析背景颜色 RGB: {RgbColor}", sourceStyle.Fill.BackgroundColor.Rgb);
                                        }
                                        // targetCell.Style.Fill.BackgroundColor.SetColor(sourceStyle.Fill.BackgroundColor.Rgb); // 旧的错误代码
                                     }
                                     
                                     targetCell.Style.HorizontalAlignment = sourceStyle.HorizontalAlignment;
                                     targetCell.Style.VerticalAlignment = sourceStyle.VerticalAlignment;
                                     targetCell.Style.WrapText = sourceStyle.WrapText;
                                     
                                     // 显式复制所有边框样式
                                     targetCell.Style.Border.Top.Style = sourceStyle.Border.Top.Style;
                                     targetCell.Style.Border.Bottom.Style = sourceStyle.Border.Bottom.Style;
                                     targetCell.Style.Border.Left.Style = sourceStyle.Border.Left.Style;
                                     targetCell.Style.Border.Right.Style = sourceStyle.Border.Right.Style;
                                     
                                     // 修改：解析颜色字符串并复制边框颜色
                                     Action<ExcelBorderItem, ExcelBorderItem> copyBorderColor = (targetBorder, sourceBorder) => {
                                         if (!string.IsNullOrEmpty(sourceBorder.Color.Rgb)) {
                                             try {
                                                 targetBorder.Color.SetColor(System.Drawing.ColorTranslator.FromHtml("#" + sourceBorder.Color.Rgb.Substring(2)));
                                             }
                                             catch (Exception ex) {
                                                _logger.LogWarning(ex, "无法解析边框颜色 RGB: {RgbColor}", sourceBorder.Color.Rgb);
                                             }
                                         }
                                     };
                                     copyBorderColor(targetCell.Style.Border.Top, sourceStyle.Border.Top);
                                     copyBorderColor(targetCell.Style.Border.Bottom, sourceStyle.Border.Bottom);
                                     copyBorderColor(targetCell.Style.Border.Left, sourceStyle.Border.Left);
                                     copyBorderColor(targetCell.Style.Border.Right, sourceStyle.Border.Right);
                                     
                                     // targetCell.Style.Border.Top.Color.SetColor(sourceStyle.Border.Top.Color.Rgb); // 旧的错误代码
                                     // targetCell.Style.Border.Bottom.Color.SetColor(sourceStyle.Border.Bottom.Color.Rgb);
                                     // targetCell.Style.Border.Left.Color.SetColor(sourceStyle.Border.Left.Color.Rgb);
                                     // targetCell.Style.Border.Right.Color.SetColor(sourceStyle.Border.Right.Color.Rgb);

                                     // 复制数字格式
                                     targetCell.Style.Numberformat.Format = sourceStyle.Numberformat.Format;
                                     
                                 }
                             }
                        } else { isFirstRow = false; }

                        string targetDisplay = resultItem.TargetName ?? "未知靶标";
                        string ctDisplay = resultItem.CtValue.HasValue ? resultItem.CtValue.Value.ToString("0.00") : (resultItem.CtValueSpecialMark ?? "-");
                        string concentrationDisplay = "-";
                        if (resultItem.Concentration.HasValue) { /* 格式化浓度 */ double conc = resultItem.Concentration.Value; if (conc >= 1e6) concentrationDisplay = $"{conc / 1e6:F2}E+06"; else if (conc >= 1e3) concentrationDisplay = $"{conc / 1e3:F2}E+03"; else if (conc > 0) concentrationDisplay = conc.ToString("F2"); else concentrationDisplay = "-"; }
                        string resultText = resultItem.DetectionResult ?? "未知";

                        _logger.LogInformation("填充阳性数据行 {Row}: 靶标={Target}, CT={CtValue}, 浓度={Concentration}, 结果={Result}", currentRow, targetDisplay, ctDisplay, concentrationDisplay, resultText);

                        // --- 填充单元格 ---
                        // 修改：确保使用正确的列索引填充
                        if (pathogenCol > 0) worksheet.Cells[currentRow, pathogenCol].Value = targetDisplay;
                        if (concentrationCol > 0) { 
                            worksheet.Cells[currentRow, concentrationCol].Value = concentrationDisplay;
                            if (concentrationDisplay != "-") { worksheet.Cells[currentRow, concentrationCol].Style.Font.Color.SetColor(Color.Red); }
                        }
                        if (ctValueCol > 0) {
                            if (resultItem.CtValue.HasValue) { worksheet.Cells[currentRow, ctValueCol].Value = resultItem.CtValue.Value; worksheet.Cells[currentRow, ctValueCol].Style.Numberformat.Format = "0.00"; } 
                            else { worksheet.Cells[currentRow, ctValueCol].Value = ctDisplay; }
                        }
                        if (resultCol > 0) {
                            worksheet.Cells[currentRow, resultCol].Value = resultText;
                            worksheet.Cells[currentRow, resultCol].Style.Font.Color.SetColor(Color.Red);
                            worksheet.Cells[currentRow, resultCol].Style.Font.Bold = true;
                        }
                    }
                }
                _logger.LogInformation("病原体列表处理完成 (使用分析结果)");
            } else {
                _logger.LogWarning("未找到合适的表头行，无法处理病原体列表 (使用分析结果)");
            }
            await Task.CompletedTask;
        }
        
        /// <summary>
        /// 辅助方法：根据占位符查找列索引 (在指定行扫描)
        /// </summary>
        private int? FindColumnByPlaceholder(ExcelWorksheet worksheet, int row, string placeholder)
        {
            if (worksheet.Dimension == null || row > worksheet.Dimension.End.Row) return null;
            for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
            {
                if (GetCellActualText(worksheet.Cells[row, col]).Contains(placeholder))
                {
                    _logger.LogDebug("通过占位符 '{Placeholder}' 在行 {Row} 找到列: {Col}", placeholder, row, col);
                    return col;
                }
            }
            return null;
        }

        /// <summary>
        /// 计算浓度值，根据分析方法中的配置
        /// </summary>
        private string CalculateConcentration(Plate plate, double? ctValue, string channel, string targetName)
        {
            if (!ctValue.HasValue || ctValue.Value <= 0) return "-";
            
            double ct = ctValue.Value;
            double baseValue = 0;
            
            try
            {
                // 默认公式参数
                double standardA = 40; // 默认标准A（截距）
                double standardB = 3.32; // 默认标准B（斜率）
                
                // 尝试从分析方法配置中获取公式参数
                if (plate.AnalysisMethodId.HasValue)
                {
                    _logger.LogInformation("尝试查找分析方法 ID={MethodId}, Name={MethodName} 中的规则", 
                        plate.AnalysisMethodId, plate.AnalysisMethod);
                    
                    // 查找此靶标/通道对应的分析规则
                    var rule = plate.AnalysisMethod != null && !string.IsNullOrEmpty(targetName) ?
                        FindAnalysisRule(plate, targetName, channel) : null;
                    
                    if (rule != null)
                    {
                        _logger.LogInformation("使用 {TargetName} 的配置公式计算浓度", targetName);
                        // 使用规则中的参数
                        standardA = rule.StandardA;
                        standardB = rule.StandardB;
                    }
                    else
                    {
                        _logger.LogWarning("未找到 {TargetName} 的分析规则，使用默认公式", targetName);
                    }
                }
                
                // 使用公式：浓度 = 10^((standardA-Ct)/standardB)
                baseValue = Math.Pow(10, (standardA - ct) / standardB);
                
                // 格式化显示
                string concValue;
                if (baseValue >= 1000000)
                {
                    // 大于等于1,000,000，使用E+06格式
                    concValue = $"{baseValue / 1000000:F2}E+06";
                }
                else if (baseValue >= 1000)
                {
                    // 大于等于1,000，使用E+03格式
                    concValue = $"{baseValue / 1000:F2}E+03";
                }
                else if (baseValue >= 100)
                {
                    // 大于等于100，保留两位小数
                    concValue = $"{baseValue:F2}";
                }
                else if (baseValue >= 10)
                {
                    // 大于等于10，保留两位小数
                    concValue = $"{baseValue:F2}";
                }
                else if (baseValue >= 1)
                {
                    // 大于等于1，保留两位小数
                    concValue = $"{baseValue:F2}";
                }
                else if (baseValue > 0)
                {
                    // 小于1，保留三位有效数字的科学计数法
                    concValue = $"{baseValue:E2}";
                }
                else
                {
                    concValue = "-";
                }
                    
                _logger.LogDebug("计算浓度: CT={CtValue}, 浓度={Concentration}", ct, concValue);
                return concValue;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "计算浓度时出错: {TargetName}, CT={CtValue}", targetName, ct);
                return "-";
            }
        }
        
        /// <summary>
        /// 查找特定靶标和通道的分析规则
        /// </summary>
        private dynamic FindAnalysisRule(Plate plate, string targetName, string channel)
        {
            try
            {
                // 注意：这里的实现是示例性的，需要根据实际的AnalysisMethod结构进行调整
                // 假设plate.AnalysisMethod中包含分析方法和规则
                if (plate.AnalysisMethodId.HasValue)
                {
                    _logger.LogInformation("尝试查找分析方法 ID={MethodId}, Name={MethodName} 中的规则", 
                        plate.AnalysisMethodId, plate.AnalysisMethod);
                    
                    // 假设规则存储在某个地方，比如板中的dynamic属性或相关对象中
                    // 这里返回一个动态对象，包含必要的参数
                    return new 
                    { 
                        StandardA = 40.0, // 默认为40
                        StandardB = 3.32, // 默认为3.32
                        // 其他参数...
                    };
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "查找分析规则时出错: {TargetName}, Channel={Channel}", targetName, channel);
            }
            
            return null;
        }

        /// <summary>
        /// 根据通道名称推断可能的病原体名称，仅在没有靶标名称时使用
        /// </summary>
        private string GetPathogenNameByChannel(string? channel)
        {
            if (string.IsNullOrEmpty(channel))
                return "未知靶标";
            
            // 这里不再硬编码返回固定值，仅当无法获得靶标名称时作为后备方案
            // 可以通过配置文件或数据库动态获取这些映射关系
            switch (channel.ToUpper())
            {
                case "FAM":
                case "VIC":
                case "ROX":
                case "CY5":
                    // 直接返回通道名称，避免使用固定映射
                    return channel;
                default:
                    return channel;
            }
        }
    }
} 