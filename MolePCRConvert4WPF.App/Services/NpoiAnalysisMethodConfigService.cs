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
            return await Task.Run(() =>
            {
                _logger.LogInformation("Attempting to load configuration from Excel file: {FilePath}", filePath);
                if (!File.Exists(filePath))
                {
                    _logger.LogError("File not found: {FilePath}", filePath);
                    throw new FileNotFoundException("指定的Excel文件未找到", filePath);
                }

                var configuration = new ObservableCollection<AnalysisMethodRule>();
                ISheet? worksheet = null;
                int headerRowIndex = -1;

                try
                {
                    using (var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) // Use FileShare.ReadWrite for potential external access
                    {
                        IWorkbook workbook;
                        if (filePath.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                        {
                            workbook = new XSSFWorkbook(fileStream);
                        }
                        else if (filePath.EndsWith(".xls", StringComparison.OrdinalIgnoreCase))
                        {
                            workbook = new HSSFWorkbook(fileStream);
                        }
                        else
                        {
                            throw new ArgumentException("不支持的文件格式，请选择 .xlsx 或 .xls 文件", nameof(filePath));
                        }

                        worksheet = workbook.GetSheetAt(0);
                        if (worksheet == null)
                        {
                            _logger.LogWarning("Excel file {FilePath} does not contain any worksheets.", filePath);
                            return configuration; // Return empty collection if no sheet exists
                        }

                        // Find header row (simple check for first row with expected headers)
                        // A more robust approach might search the first few rows
                        IRow? potentialHeaderRow = worksheet.GetRow(0);
                        if (potentialHeaderRow != null && IsHeaderRow(potentialHeaderRow))
                        {
                            headerRowIndex = 0;
                        }
                        else
                        {
                             _logger.LogWarning("Could not find expected header row in {FilePath}", filePath);
                            // Optionally: Try row 1 as header? Or throw an error?
                             return configuration; // Cannot proceed without headers
                        }
                    }

                    int rowCount = worksheet.LastRowNum + 1;
                     _logger.LogDebug("Worksheet '{SheetName}' has {RowCount} rows (including header).", worksheet.SheetName, rowCount);

                    // Start reading from the row after the header
                    for (int rowIndex = headerRowIndex + 1; rowIndex < rowCount; rowIndex++)
                    {
                        IRow? row = worksheet.GetRow(rowIndex);
                        if (row == null || row.Cells.All(d => d.CellType == CellType.Blank)) // Skip empty rows
                        {
                            continue;
                        }

                        try
                        {
                             var rule = new AnalysisMethodRule
                             {
                                // Use helper to safely get cell values
                                Index = int.TryParse(GetCellValueAsString(row.GetCell(0)), out int index) ? index : rowIndex, // Use row index if '序号' is invalid or missing
                                WellPosition = GetCellValueAsString(row.GetCell(1)),
                                WellPositionPattern = GetCellValueAsString(row.GetCell(1)), // 确保WellPositionPattern设置与WellPosition相同
                                Channel = GetCellValueAsString(row.GetCell(2)),
                                SpeciesName = GetCellValueAsString(row.GetCell(3)),
                                TargetName = GetCellValueAsString(row.GetCell(3)), // 确保TargetName设置与SpeciesName相同
                                JudgeFormula = GetCellValueAsString(row.GetCell(4)),
                                ConcentrationFormula = GetCellValueAsString(row.GetCell(5))
                             };
                            configuration.Add(rule);
                        }
                        catch (Exception exRow)
                        {
                            _logger.LogWarning(exRow, "Error reading row {RowIndex} in file {FilePath}. Skipping row.", rowIndex + 1, filePath);
                            // Decide whether to continue or throw
                        }
                    }
                    _logger.LogInformation("Successfully loaded {RuleCount} rules from {FilePath}", configuration.Count, filePath);
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Failed to load configuration from Excel file: {FilePath}", filePath);
                    // Wrap NPOI exceptions or other file access exceptions
                    throw new Exception($"读取Excel文件 '{Path.GetFileName(filePath)}' 时出错: {ex.Message}", ex);
                }

                return configuration;
            });
        }

        public async Task SaveConfigurationAsync(string filePath, ObservableCollection<AnalysisMethodRule> configuration)
        {
             if (configuration == null) throw new ArgumentNullException(nameof(configuration));
             
            await Task.Run(() =>
            {
                 _logger.LogInformation("Attempting to save {RuleCount} rules to Excel file: {FilePath}", configuration.Count, filePath);
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
                     throw new ArgumentException("不支持的文件格式，请使用 .xlsx 或 .ls 文件", nameof(filePath));
                }

                ISheet sheet = workbook.CreateSheet("分析方法"); // Or use a configurable sheet name

                try
                {
                    // Create Header Row
                    IRow headerRow = sheet.CreateRow(0);
                    ICellStyle headerStyle = workbook.CreateCellStyle();
                    IFont headerFont = workbook.CreateFont();
                    headerFont.IsBold = true;
                    // Optional: Set font size, color etc.
                    headerStyle.SetFont(headerFont);

                    for (int i = 0; i < _expectedHeaders.Length; i++)
                    {
                        ICell cell = headerRow.CreateCell(i);
                        cell.SetCellValue(_expectedHeaders[i]);
                        cell.CellStyle = headerStyle;
                    }

                    // Write Data Rows
                    int rowIndex = 1;
                    foreach (var rule in configuration)
                    {
                        IRow dataRow = sheet.CreateRow(rowIndex++);
                        dataRow.CreateCell(0).SetCellValue(rule.Index); // Assuming Index is numeric
                        dataRow.CreateCell(1).SetCellValue(rule.WellPosition ?? string.Empty);
                        dataRow.CreateCell(2).SetCellValue(rule.Channel ?? string.Empty);
                        dataRow.CreateCell(3).SetCellValue(rule.SpeciesName ?? string.Empty);
                        dataRow.CreateCell(4).SetCellValue(rule.JudgeFormula ?? string.Empty);
                        dataRow.CreateCell(5).SetCellValue(rule.ConcentrationFormula ?? string.Empty);
                    }

                    // Auto-size columns (can be slow for large files, consider setting fixed widths)
                     for (int i = 0; i < _expectedHeaders.Length; i++)
                     {
                         sheet.AutoSizeColumn(i);
                     }

                    // Write to file
                    // Ensure directory exists
                    string? directory = Path.GetDirectoryName(filePath);
                    if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                    {
                        Directory.CreateDirectory(directory);
                         _logger.LogInformation("Created directory: {DirectoryPath}", directory);
                    }

                    using (var fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                    {
                        workbook.Write(fileStream);
                    }
                     _logger.LogInformation("Successfully saved configuration to {FilePath}", filePath);
                }
                catch (Exception ex)
                {
                     _logger.LogError(ex, "Failed to save configuration to Excel file: {FilePath}", filePath);
                    // Consider specific exception handling for file access vs NPOI issues
                    throw new Exception($"保存Excel文件 '{Path.GetFileName(filePath)}' 时出错: {ex.Message}", ex);
                }
                finally // Ensure workbook resources are closed if using HSSFWorkbook
                {
                     if (!useXlsx)
                     {
                         // HSSFWorkbook might need explicit close/dispose, though using statement should handle it.
                         // workbook.Close(); // Generally not needed with 'using' on stream
                     }
                }
            });
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
                    
                    // Optional: Add sample/default data rows like in the old project
                    // Example:
                     /*
                     IRow row1 = sheet.CreateRow(1);
                     row1.CreateCell(0).SetCellValue(1);
                     row1.CreateCell(1).SetCellValue("A1");
                     row1.CreateCell(2).SetCellValue("FAM");
                     row1.CreateCell(3).SetCellValue("示例病毒");
                     row1.CreateCell(4).SetCellValue("[CT]<38");
                     row1.CreateCell(5).SetCellValue(""); 
                     */

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
    }
} 