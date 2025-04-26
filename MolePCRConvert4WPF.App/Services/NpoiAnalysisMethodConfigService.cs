using Microsoft.Extensions.Logging;
using MolePCRConvert4WPF.Core.Interfaces;
using MolePCRConvert4WPF.Core.Models;
using NPOI.HSSF.UserModel;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MolePCRConvert4WPF.App.Services // Or Infrastructure if you have one
{
    /// <summary>
    /// Implementation of IAnalysisMethodConfigService using the NPOI library.
    /// </summary>
    public class NpoiAnalysisMethodConfigService : IAnalysisMethodConfigService
    {
        private readonly ILogger<NpoiAnalysisMethodConfigService> _logger;

        // Define expected header names (adjust if needed)
        private readonly string[] _expectedHeaders = { "序号", "孔位", "荧光通道", "种名", "阳性判定公式", "浓度计算公式" };

        public NpoiAnalysisMethodConfigService(ILogger<NpoiAnalysisMethodConfigService> logger)
        {
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            // NPOI License (if using a version that requires it for commercial use)
            // NPOI.NpoiLicense.LicenseKey = "your-license-key"; 
        }

        public async Task<ObservableCollection<AnalysisMethodRule>> LoadConfigurationAsync(string filePath)
        {
            if (!File.Exists(filePath))
            {
                _logger.LogError("文件不存在: {FilePath}", filePath);
                throw new FileNotFoundException("指定的Excel文件未找到", filePath);
            }
            
            // 为确保读取最新内容，先清理文件系统缓存
            GC.Collect();
            GC.WaitForPendingFinalizers();
            
            // 确保不使用缓存，强制刷新文件内容
            _logger.LogInformation("从Excel文件加载配置数据: {FilePath}", filePath);
            
            // 获取文件信息，确保读取的是最新文件
            FileInfo fileInfo = new FileInfo(filePath);
            _logger.LogInformation("文件大小: {Size} 字节, 最后修改时间: {LastWriteTime}", 
                                   fileInfo.Length, fileInfo.LastWriteTime);
            
            var configuration = new ObservableCollection<AnalysisMethodRule>();
            
            // 允许UI线程继续运行
            await Task.Delay(10);
            
            try
            {
                // 使用FileShare.ReadWrite允许其他进程同时访问文件
                using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    IWorkbook workbook;
                    if (filePath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                    {
                        workbook = new XSSFWorkbook(stream);
                    }
                    else if (filePath.EndsWith(".xls", StringComparison.OrdinalIgnoreCase))
                    {
                        workbook = new HSSFWorkbook(stream);
                    }
                    else
                    {
                        throw new ArgumentException("不支持的文件格式，请选择 .xlsx 或 .xls 文件", nameof(filePath));
                    }

                    try
                    {
                        var worksheet = workbook.GetSheetAt(0);
                        if (worksheet == null)
                        {
                            _logger.LogWarning("Excel文件 {FilePath} 不包含任何工作表", filePath);
                            workbook.Close();
                            return configuration;
                        }

                        // 查找标题行
                        int headerRowIndex = -1;
                        IRow potentialHeaderRow = worksheet.GetRow(0);
                        if (potentialHeaderRow != null && IsHeaderRow(potentialHeaderRow))
                        {
                            headerRowIndex = 0;
                        }
                        else
                        {
                            _logger.LogWarning("在 {FilePath} 中找不到预期的标题行", filePath);
                            workbook.Close();
                            return configuration;
                        }

                        int rowCount = worksheet.LastRowNum + 1;
                        _logger.LogDebug("工作表 '{SheetName}' 共有 {RowCount} 行(包括标题)", worksheet.SheetName, rowCount);

                        // 从标题行后的行开始读取
                        for (int rowIndex = headerRowIndex + 1; rowIndex < rowCount; rowIndex++)
                        {
                            IRow row = worksheet.GetRow(rowIndex);
                            if (row == null || IsEmptyRow(row))
                            {
                                continue;
                            }

                            try
                            {
                                var rule = new AnalysisMethodRule
                                {
                                    Index = int.TryParse(GetCellValueAsString(row.GetCell(0)), out int index) ? index : rowIndex,
                                    WellPosition = GetCellValueAsString(row.GetCell(1)),
                                    WellPositionPattern = GetCellValueAsString(row.GetCell(1)),
                                    Channel = GetCellValueAsString(row.GetCell(2)),
                                    SpeciesName = GetCellValueAsString(row.GetCell(3)),
                                    TargetName = GetCellValueAsString(row.GetCell(3)),
                                    JudgeFormula = GetCellValueAsString(row.GetCell(4)),
                                    ConcentrationFormula = GetCellValueAsString(row.GetCell(5))
                                };
                                configuration.Add(rule);
                                
                                _logger.LogDebug("加载行 {Row}: 序号={Index}, 孔位={Well}, 通道={Channel}, 种名={Species}",
                                                rowIndex, rule.Index, rule.WellPosition, rule.Channel, rule.SpeciesName);
                            }
                            catch (Exception exRow)
                            {
                                _logger.LogWarning(exRow, "读取第 {RowIndex} 行时出错, 跳过该行", rowIndex + 1);
                            }
                        }
                        
                        // 记录一些关键数据以便调试
                        if (configuration.Count > 0)
                        {
                            var firstRow = configuration[0];
                            _logger.LogInformation("首行数据: 序号={Index}, 孔位={Well}, 种名={Species}", 
                                                 firstRow.Index, firstRow.WellPosition, firstRow.SpeciesName);
                            
                            if (configuration.Count > 1)
                            {
                                var lastRow = configuration[configuration.Count - 1];
                                _logger.LogInformation("末行数据: 序号={Index}, 孔位={Well}, 种名={Species}", 
                                                     lastRow.Index, lastRow.WellPosition, lastRow.SpeciesName);
                            }
                        }
                    }
                    finally
                    {
                        // 确保工作簿被关闭
                        workbook.Close();
                    }
                }
                
                _logger.LogInformation("成功从 {FilePath} 加载了 {Count} 条规则", filePath, configuration.Count);

                // === 日志记录：服务层加载后的数据 ===
                _logger.LogInformation("--- 服务层加载后 configuration 数据快照 ---");
                for (int i = 0; i < Math.Min(5, configuration.Count); i++)
                {
                    var rule = configuration[i];
                    _logger.LogInformation($"  Index={rule.Index}, Well={rule.WellPosition}, Channel={rule.Channel}, Species={rule.SpeciesName}, Judge='{rule.JudgeFormula}', Conc='{rule.ConcentrationFormula}'");
                }
                if (configuration.Count > 10) _logger.LogInformation("  ...");
                for (int i = Math.Max(0, configuration.Count - 5); i < configuration.Count; i++)
                {
                     var rule = configuration[i];
                    _logger.LogInformation($"  Index={rule.Index}, Well={rule.WellPosition}, Channel={rule.Channel}, Species={rule.SpeciesName}, Judge='{rule.JudgeFormula}', Conc='{rule.ConcentrationFormula}'");
                }
                _logger.LogInformation("-----------------------------------------");
                // ====================================

            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "从Excel文件加载配置失败: {FilePath}", filePath);
                throw new Exception($"读取Excel文件 '{Path.GetFileName(filePath)}' 时出错: {ex.Message}", ex);
            }

            return configuration;
        }

        public async Task SaveConfigurationAsync(string filePath, ObservableCollection<AnalysisMethodRule> configuration)
        {
            if (configuration == null) throw new ArgumentNullException(nameof(configuration));
            
            // 使用Task.Delay允许UI线程继续运行
            await Task.Delay(1);
            
            _logger.LogInformation("开始保存数据到Excel文件: {FilePath}", filePath);
            
            // 创建备份
            string backupFilePath = string.Empty;
            if (File.Exists(filePath))
            {
                try 
                {
                    backupFilePath = filePath + ".bak";
                    File.Copy(filePath, backupFilePath, true);
                    _logger.LogInformation("已创建备份文件: {BackupPath}", backupFilePath);
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "创建备份文件失败，继续执行保存操作");
                }
            }
            
            try
            {
                // 创建新工作簿
                IWorkbook workbook;
                if (filePath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                {
                    workbook = new XSSFWorkbook();
                }
                else if (filePath.EndsWith(".xls", StringComparison.OrdinalIgnoreCase))
                {
                    workbook = new HSSFWorkbook();
                }
                else
                {
                    throw new ArgumentException("不支持的文件格式，请使用 .xlsx 或 .xls 文件", nameof(filePath));
                }

                // 创建工作表
                ISheet sheet = workbook.CreateSheet("分析方法");
                
                // 创建标题样式
                ICellStyle headerStyle = workbook.CreateCellStyle();
                IFont headerFont = workbook.CreateFont();
                headerFont.IsBold = true;
                headerStyle.SetFont(headerFont);
                
                // 创建标题行
                IRow headerRow = sheet.CreateRow(0);
                for (int i = 0; i < _expectedHeaders.Length; i++)
                {
                    ICell cell = headerRow.CreateCell(i);
                    cell.SetCellValue(_expectedHeaders[i]);
                    cell.CellStyle = headerStyle;
                }
                
                // 写入数据行
                int rowIndex = 1;
                foreach (var rule in configuration)
                {
                    IRow dataRow = sheet.CreateRow(rowIndex++);
                    dataRow.CreateCell(0).SetCellValue(rule.Index); 
                    dataRow.CreateCell(1).SetCellValue(rule.WellPosition ?? string.Empty);
                    dataRow.CreateCell(2).SetCellValue(rule.Channel ?? string.Empty);
                    dataRow.CreateCell(3).SetCellValue(rule.SpeciesName ?? string.Empty);
                    dataRow.CreateCell(4).SetCellValue(rule.JudgeFormula ?? string.Empty);
                    dataRow.CreateCell(5).SetCellValue(rule.ConcentrationFormula ?? string.Empty);
                }
                
                // 调整列宽
                for (int i = 0; i < _expectedHeaders.Length; i++)
                {
                    sheet.AutoSizeColumn(i);
                }
                
                // 强制删除目标文件（如果存在）
                if (File.Exists(filePath))
                {
                    // 强制释放文件句柄
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                    
                    // 尝试删除文件
                    try 
                    {
                        File.Delete(filePath);
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError(ex, "无法删除原文件，尝试直接覆盖: {FilePath}", filePath);
                    }
                }
                
                // 修复：使用内存流先将数据写入内存，避免直接使用文件流可能导致的资源争用问题
                byte[] fileContent;
                using (MemoryStream ms = new MemoryStream())
                {
                    workbook.Write(ms);
                    fileContent = ms.ToArray();
                }
                
                // 关闭工作簿释放资源
                workbook.Close();
                workbook = null;
                
                // 强制释放所有资源
                GC.Collect();
                GC.WaitForPendingFinalizers();
                
                // === 日志记录：服务层保存前的数据 ===
                _logger.LogInformation("--- 服务层保存前 configuration 数据快照 ---");
                 for (int i = 0; i < Math.Min(5, configuration.Count); i++)
                {
                    var rule = configuration[i];
                    _logger.LogInformation($"  Index={rule.Index}, Well={rule.WellPosition}, Channel={rule.Channel}, Species={rule.SpeciesName}, Judge='{rule.JudgeFormula}', Conc='{rule.ConcentrationFormula}'");
                }
                if (configuration.Count > 10) _logger.LogInformation("  ...");
                for (int i = Math.Max(0, configuration.Count - 5); i < configuration.Count; i++)
                {
                     var rule = configuration[i];
                    _logger.LogInformation($"  Index={rule.Index}, Well={rule.WellPosition}, Channel={rule.Channel}, Species={rule.SpeciesName}, Judge='{rule.JudgeFormula}', Conc='{rule.ConcentrationFormula}'");
                }
                 _logger.LogInformation("-----------------------------------------");
                // ====================================

                // 将内存中的数据一次性写入文件
                File.WriteAllBytes(filePath, fileContent);
                
                // 验证文件是否成功写入
                if (File.Exists(filePath))
                {
                    FileInfo fileInfo = new FileInfo(filePath);
                    if (fileInfo.Length > 0)
                    {
                        _logger.LogInformation("成功保存文件，大小: {Size} 字节", fileInfo.Length);
                        
                        // 删除备份文件
                        if (!string.IsNullOrEmpty(backupFilePath) && File.Exists(backupFilePath))
                        {
                            try { File.Delete(backupFilePath); }
                            catch { /* 忽略删除备份时的错误 */ }
                        }
                    }
                    else
                    {
                        _logger.LogWarning("文件保存成功但大小为0，可能存在问题");
                        throw new IOException("保存的文件大小为0，数据可能未正确写入");
                    }
                }
                else
                {
                    _logger.LogError("文件保存后不存在，保存失败");
                    throw new IOException("文件保存后不存在，保存失败");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "保存Excel文件时出错: {FilePath}", filePath);
                
                // 恢复备份
                if (!string.IsNullOrEmpty(backupFilePath) && File.Exists(backupFilePath))
                {
                    try 
                    { 
                        File.Copy(backupFilePath, filePath, true);
                        _logger.LogInformation("已恢复备份文件");
                    }
                    catch { /* 忽略恢复备份时的错误 */ }
                }
                
                throw new Exception($"保存Excel文件 '{Path.GetFileName(filePath)}' 时出错: {ex.Message}", ex);
            }
        }

        public async Task CreateNewConfigurationFileAsync(string filePath)
        {
            await Task.Run(() =>
            {
                 _logger.LogInformation("Attempting to create new configuration file: {FilePath}", filePath);
                if (File.Exists(filePath))
                {
                    // Consider adding an 'overwrite' parameter or throwing
                     _logger.LogWarning("File already exists, creation skipped: {FilePath}", filePath);
                     throw new IOException($"文件 '{Path.GetFileName(filePath)}' 已存在.");
                }
                
                IWorkbook workbook;
                bool useXlsx = filePath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase);

                 if (useXlsx)
                {
                    workbook = new XSSFWorkbook();
                }
                else if (filePath.EndsWith(".xls", StringComparison.OrdinalIgnoreCase))
                {
                    workbook = new HSSFWorkbook();
                }
                else
                {
                     throw new ArgumentException("不支持的文件格式，请使用 .xlsx 或 .xls 文件", nameof(filePath));
                }
                
                ISheet sheet = workbook.CreateSheet("分析方法");

                try
                {
                    // Create Header Row (same as in SaveConfigurationAsync)
                    IRow headerRow = sheet.CreateRow(0);
                    ICellStyle headerStyle = workbook.CreateCellStyle();
                    IFont headerFont = workbook.CreateFont();
                    headerFont.IsBold = true;
                    headerStyle.SetFont(headerFont);
                    for (int i = 0; i < _expectedHeaders.Length; i++)
                    {
                        ICell cell = headerRow.CreateCell(i);
                        cell.SetCellValue(_expectedHeaders[i]);
                        cell.CellStyle = headerStyle;
                    }
                    
                    // 添加一些默认的示例数据行，方便用户编辑
                    AddDefaultRow(sheet, 1, "A1", "FAM", "呼吸道合胞病毒", "[CT]<36", "");
                    AddDefaultRow(sheet, 2, "A1", "VIC", "乙型流感病毒", "[CT]<36", "");
                    AddDefaultRow(sheet, 3, "A1", "ROX", "腺病毒", "[CT]<36", "");
                    AddDefaultRow(sheet, 4, "A1", "CY5", "甲型流感病毒", "[CT]<36", "");

                    // Auto-size columns
                     for (int i = 0; i < _expectedHeaders.Length; i++)
                     {
                         sheet.AutoSizeColumn(i);
                     }

                    // Write to file
                     string? directory = Path.GetDirectoryName(filePath);
                    if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                    {
                        Directory.CreateDirectory(directory);
                         _logger.LogInformation("Created directory: {DirectoryPath}", directory);
                    }
                    
                    using (var fileStream = new FileStream(filePath, FileMode.CreateNew, FileAccess.Write)) // Use CreateNew to avoid accidental overwrite
                    {
                        workbook.Write(fileStream);
                    }
                     _logger.LogInformation("Successfully created new configuration file: {FilePath}", filePath);
                }
                catch(Exception ex)
                {
                     _logger.LogError(ex, "Failed to create new configuration file: {FilePath}", filePath);
                    throw new Exception($"创建新的Excel文件 '{Path.GetFileName(filePath)}' 时出错: {ex.Message}", ex);
                }
            });
        }

        // 添加默认数据行的辅助方法
        private void AddDefaultRow(ISheet sheet, int rowIndex, string wellPosition, string channel, string speciesName, string judgeFormula, string concentrationFormula)
        {
            IRow row = sheet.CreateRow(rowIndex);
            row.CreateCell(0).SetCellValue(rowIndex);
            row.CreateCell(1).SetCellValue(wellPosition);
            row.CreateCell(2).SetCellValue(channel);
            row.CreateCell(3).SetCellValue(speciesName);
            row.CreateCell(4).SetCellValue(judgeFormula);
            row.CreateCell(5).SetCellValue(concentrationFormula);
        }
        
        // --- Helper Methods --- 

        /// <summary>
        /// Safely gets the string value of a cell, handling different cell types.
        /// </summary>
        private string GetCellValueAsString(ICell? cell)
        {
            if (cell == null) return string.Empty;

            switch (cell.CellType)
            {
                case CellType.String:
                    return cell.StringCellValue;
                case CellType.Numeric:
                    // Handle date/time if necessary, otherwise format as string
                    if (DateUtil.IsCellDateFormatted(cell))
                    {
                        try 
                        {
                            // Attempt to get DateTime and format it
                            DateTime? nullableDateValue = cell.DateCellValue; // Get as nullable DateTime?
                            if (nullableDateValue.HasValue)
                            {
                                return nullableDateValue.Value.ToString("yyyy-MM-dd HH:mm:ss"); // Access .Value if not null
                            }
                            else
                            {
                                // Handle case where DateCellValue is null despite IsCellDateFormatted being true (unlikely but possible)
                                _logger?.LogWarning("Cell {Address} is date formatted but DateCellValue is null, falling back to numeric string.", cell.Address);
                                return cell.NumericCellValue.ToString(System.Globalization.CultureInfo.InvariantCulture);
                            }
                        } 
                        catch (Exception ex) // Catch potential exceptions during DateCellValue access or formatting
                        { 
                            _logger?.LogWarning(ex, "Could not format numeric cell {Address} as date, falling back to numeric string.", cell.Address);
                            // Fallback to numeric string representation if date formatting fails
                            return cell.NumericCellValue.ToString(System.Globalization.CultureInfo.InvariantCulture); 
                        }
                    }
                    else
                    {
                        // Use InvariantCulture to avoid issues with decimal separators
                        return cell.NumericCellValue.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    }
                case CellType.Boolean:
                    return cell.BooleanCellValue.ToString();
                case CellType.Formula:
                    // Try evaluating the formula, fallback to cached value or formula string
                    try
                    {
                        // Use FormulaEvaluator for complex formulas if needed
                        // IFormulaEvaluator evaluator = cell.Sheet.Workbook.GetCreationHelper().CreateFormulaEvaluator();
                        // return evaluator.Evaluate(cell).FormatAsString(); // More robust evaluation
                        return cell.StringCellValue; // Often stores the result as string for simple formulas
                    }
                    catch
                    {
                        try { return cell.NumericCellValue.ToString(); } // Try numeric cached value
                        catch { return cell.CellFormula; } // Fallback to the formula itself
                    }
                case CellType.Blank:
                    return string.Empty;
                case CellType.Error:
                     return FormulaError.ForInt(cell.ErrorCellValue).String;
                case CellType.Unknown:
                default:
                    return string.Empty;
            }
        }
        
        /// <summary>
        /// Checks if a row likely contains the expected headers.
        /// </summary>
        private bool IsHeaderRow(IRow headerRow)
        {
            if (headerRow.LastCellNum < _expectedHeaders.Length) return false;
            for (int i = 0; i < _expectedHeaders.Length; i++)
            {
                ICell? cell = headerRow.GetCell(i);
                if (cell == null || !GetCellValueAsString(cell).Equals(_expectedHeaders[i], StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }
            }
            return true;
        }

        private bool IsEmptyRow(IRow row)
        {
            if (row == null) return true;
            
            // 检查所有单元格是否为空
            for (int i = 0; i < 6; i++) // 只检查我们关心的6列
            {
                ICell cell = row.GetCell(i);
                if (cell != null && cell.CellType != CellType.Blank && !string.IsNullOrWhiteSpace(GetCellValueAsString(cell)))
                {
                    return false; // 找到非空单元格
                }
            }
            
            return true; // 所有单元格都为空
        }
    }
} 