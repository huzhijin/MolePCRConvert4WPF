using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using MolePCRConvert4WPF.Core.Enums;
using MolePCRConvert4WPF.Core.Interfaces;
using MolePCRConvert4WPF.Core.Models;
// using OfficeOpenXml; // Remove EPPlus using
using NPOI.SS.UserModel; // Add NPOI using
using NPOI.HSSF.UserModel; // Add NPOI specific for .xls
using NPOI.XSSF.UserModel; // Add NPOI specific for .xlsx
using System.Globalization;
using System.Text.RegularExpressions;

namespace MolePCRConvert4WPF.Infrastructure.FileHandlers
{
    /// <summary>
    /// Excel文件处理器 (使用 NPOI)
    /// </summary>
    public class ExcelFileHandler : BaseFileHandler
    {
        private readonly ILogger<ExcelFileHandler>? _logger;

        // Static constructor for EPPlus License context - Remove if EPPlus is fully removed
        // static ExcelFileHandler()
        // {
        //      ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        // }
        
        public ExcelFileHandler(ILogger<ExcelFileHandler>? logger = null)
        {
            _logger = logger;
        }
        
        /// <summary>
        /// 使用 NPOI 安全方式打开 Excel 工作簿 (.xls or .xlsx)
        /// </summary>
        private IWorkbook OpenWorkbookSafely(string filePath)
        {
            try
            {
                // Use FileShare.ReadWrite to potentially allow access even if the file is open elsewhere
                 using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                 {
                     // WorkbookFactory automatically detects .xls or .xlsx
                     return WorkbookFactory.Create(stream);
                 }
            }
            catch (FileNotFoundException fnfEx)
            {
                _logger?.LogError(fnfEx, "File not found: {FilePath}", filePath);
                throw; // Re-throw specific exception
            }
            catch (IOException ioEx) 
            {
                // Catch potential sharing violations or other IO errors
                 _logger?.LogError(ioEx, "IO Error opening file: {FilePath}", filePath);
                 throw new IOException($"无法打开文件 '{Path.GetFileName(filePath)}'。可能文件已损坏或被其他程序占用。", ioEx);
            }
            catch (Exception ex) 
            {
                 // Catch other potential NPOI errors during opening
                 _logger?.LogError(ex, "Error opening Excel file with NPOI: {FilePath}", filePath);
                 throw new InvalidDataException($"使用 NPOI 打开 Excel 文件时出错: {ex.Message}", ex);
            }
        }
        
        /// <summary>
        /// 检测仪器类型 (基于文件内容标记) - 使用 NPOI
        /// </summary>
        public override async Task<InstrumentType> DetectInstrumentTypeAsync(string filePath)
        {
            _logger?.LogInformation($"开始检测仪器类型 (NPOI): {filePath}");
            FileType fileType = GetFileType(filePath);
            
            if (fileType != FileType.Excel && fileType != FileType.ExcelLegacy)
            {
                _logger?.LogWarning("文件类型不是Excel，无法检测仪器类型");
                return InstrumentType.Unknown;
            }
            
            InstrumentType detectedType = InstrumentType.Unknown;
            
            await Task.Run(() =>
            {
                IWorkbook? workbook = null;
                try
                {
                    workbook = OpenWorkbookSafely(filePath);
                    if (workbook.NumberOfSheets == 0)
                    {
                         _logger?.LogWarning("Excel文件不包含任何工作表 (NPOI)");
                         return;
                    }
                    ISheet sheet = workbook.GetSheetAt(0); // Use the first sheet (index 0)
                    
                    // Determine the range to scan (e.g., first 11 rows (0-10) and 11 columns (0-10))
                    int maxScanRow = Math.Min(10, sheet.LastRowNum); // NPOI LastRowNum is 0-based index of last row
                    int maxScanCol = 10; // Scan first 11 columns (index 0-10)

                     _logger?.LogDebug($"扫描范围 (NPOI): Rows=0-{maxScanRow}, Cols=0-{maxScanCol}");

                    // Find specific markers within the scan range
                    for (int r = 0; r <= maxScanRow; r++)
                    {
                         IRow? row = sheet.GetRow(r);
                         if (row == null) continue;
                         int currentMaxCol = Math.Min(maxScanCol, row.LastCellNum); // LastCellNum is 1-based count

                        for (int c = 0; c <= currentMaxCol; c++)
                        {
                            ICell? cell = row.GetCell(c); // NPOI cell access is 0-based
                            if (cell == null) continue;
                            
                            // Attempt to get string value robustly
                            string cellValue = string.Empty;
                            try { cellValue = cell.ToString()?.Trim() ?? string.Empty; } catch { /* Ignore potential errors during ToString */ }
                            
                            if (string.IsNullOrEmpty(cellValue)) continue;
                            
                             _logger?.LogTrace($"检查单元格 [{r},{c}]: '{cellValue}'");

                            // Simplified detection logic (needs refinement based on actual file formats)
                            if (cellValue.Contains("7500", StringComparison.OrdinalIgnoreCase)) { detectedType = InstrumentType.ABI7500; return; }
                            if (cellValue.Contains("SLAN-96S", StringComparison.OrdinalIgnoreCase)) { detectedType = InstrumentType.SLAN96S; return; }
                            if (cellValue.Contains("SLAN-96P", StringComparison.OrdinalIgnoreCase)) { detectedType = InstrumentType.SLAN96P; return; }
                            if (cellValue.Contains("QuantStudio", StringComparison.OrdinalIgnoreCase) || cellValue.Contains("QS5")) { detectedType = InstrumentType.ABIQ5; return; }
                            if (cellValue.Contains("MA600", StringComparison.OrdinalIgnoreCase)) { detectedType = InstrumentType.MA600; return; }
                            if (cellValue.Contains("CFX96", StringComparison.OrdinalIgnoreCase))
                            {
                                detectedType = cellValue.Contains("Deep Well", StringComparison.OrdinalIgnoreCase) ? InstrumentType.CFX96DeepWell : InstrumentType.CFX96;
                                return;
                            } 
                        }
                    }
                    _logger?.LogInformation("未找到特定仪器标记，默认使用 ABI7500 (可配置)");
                    detectedType = InstrumentType.ABI7500; 
                }
                catch (Exception ex)
                {
                    _logger?.LogError(ex, $"检测仪器类型时出错 (NPOI): {filePath}");
                    detectedType = InstrumentType.Unknown; 
                }
                finally
                {
                     // NPOI IWorkbook doesn't implement IDisposable directly in the interface,
                     // but specific implementations (HSSFWorkbook, XSSFWorkbook) do.
                     // However, WorkbookFactory might return a wrapper. Best practice is to close if possible.
                    (workbook as IDisposable)?.Dispose();
                    // Or if using streams, ensure the stream is disposed (handled by using statement for stream)
                }
            }).ConfigureAwait(false); 

            _logger?.LogInformation($"检测到的仪器类型: {detectedType}");
            return detectedType;
        }
        
        /// <summary>
        /// 读取PCR数据 (使用 NPOI)
        /// </summary>
        public override async Task<Plate?> ReadPcrDataAsync(string filePath, InstrumentType instrumentType, Guid plateId)
        {
            _logger?.LogInformation($"开始读取PCR数据 (NPOI): 文件={filePath}, 仪器类型={instrumentType}");
            var plate = CreateNewPlate(Path.GetFileNameWithoutExtension(filePath), instrumentType, filePath, plateId);
            List<WellLayout>? parsedWells = null; // Initialize as nullable
            
            await Task.Run(() => 
            {
                 IWorkbook? workbook = null;
                 try
                 {
                    workbook = OpenWorkbookSafely(filePath);
                    if (workbook.NumberOfSheets == 0)
                    {
                         _logger?.LogError("Excel文件中找不到工作表 (NPOI)");
                         throw new InvalidDataException("Excel文件不包含任何工作表");
                    }
                    
                    // Try to find sheet named "Results", otherwise use the first sheet
                    ISheet? sheet = workbook.GetSheet("Results") ?? workbook.GetSheetAt(0); 
                    if (sheet == null)
                    {
                         _logger?.LogError("无法找到有效的 Excel 工作表 (NPOI)");
                         throw new InvalidDataException("无法找到有效的 Excel 工作表");
                    }
                     _logger?.LogInformation($"使用工作表: {sheet.SheetName}");

                    // Use NPOI specific parsing logic
                    switch (instrumentType)
                    {
                         case InstrumentType.ABI7500:
                             parsedWells = ParseABI7500Data_NPOI(sheet, plate); // Call NPOI version
                             break;
                        case InstrumentType.SLAN96S:
                        case InstrumentType.SLAN96P:
                             parsedWells = ParseSLANData_NPOI(sheet, plate);
                             break;
                        case InstrumentType.MA600:
                             parsedWells = ParseMA6000Data_NPOI(sheet, plate);
                             break;
                        case InstrumentType.ABIQ5:
                             parsedWells = ParseABIQ5Data_NPOI(sheet, plate);
                             break;
                        case InstrumentType.CFX96:
                        case InstrumentType.CFX96DeepWell:
                             parsedWells = ParseCFX96Data_NPOI(sheet, plate);
                             break;
                        // ... other cases
                        default:
                             _logger?.LogWarning($"仪器类型 {instrumentType} 尚无特定 NPOI 解析器，将返回空列表。");
                             parsedWells = new List<WellLayout>(); // Return empty list
                             break;
                    }
                    // No need to return inside try, assigned to outer scope variable 'parsedWells'
                 }
                 catch(Exception ex)
                 {
                      _logger?.LogError(ex, $"读取PCR数据时出错 (NPOI): {filePath}");
                      parsedWells = null; // Ensure parsedWells is null on error
                 }
                 finally
                 {
                     (workbook as IDisposable)?.Dispose(); 
                 }
            }).ConfigureAwait(false);

            // Check the result after Task.Run completes
            if (parsedWells != null) 
            {
                 plate.WellLayouts = parsedWells;
                 _logger?.LogInformation($"读取完成 (NPOI)，孔位数: {plate.WellLayouts.Count}");
                 return plate; // Return the populated plate
            }
            else
            {
                 _logger?.LogError($"读取PCR数据失败 (NPOI): {filePath}");
                 return null; // Return null if reading failed
            }
        }
        
        /// <summary>
        /// 校验文件格式 (Placeholder)
        /// </summary>
        public override Task<bool> ValidateFileFormatAsync(string filePath, InstrumentType instrumentType)
        {
            _logger?.LogInformation($"开始校验文件格式: {filePath}, 仪器类型={instrumentType}");
            try
            {
                // Basic check: file exists and is Excel
                FileType fileType = GetFileType(filePath);
                if (!File.Exists(filePath) || (fileType != FileType.Excel && fileType != FileType.ExcelLegacy))
                {
                    _logger?.LogWarning("文件不存在或不是Excel格式");
                    return Task.FromResult(false);
                }
                
                // TODO: Add more specific validation based on instrument type if needed
                // e.g., check for specific sheet names, header rows, etc.
                
                 _logger?.LogInformation("文件格式校验通过 (基础检查)");
                return Task.FromResult(true);
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "校验文件格式时出错");
                return Task.FromResult(false);
            }
        }

        // --- Specific NPOI Parsing Methods --- 
        
        private List<WellLayout> ParseABI7500Data_NPOI(ISheet sheet, Plate plate)
        {
            _logger?.LogInformation($"开始解析 ABI7500 数据 (NPOI) from sheet: {sheet.SheetName}");
            var parsedWells = new List<WellLayout>();

            // --- Find Header Row and Column Indices --- Find Header Row and Column Indices ---
            // NPOI uses 0-based indexing
            int headerRowIndex = 7; // Default based on Excel row 8
            int dataStartRowIndex = 8; // Default based on Excel row 9
            int wellColIndex = 0;      // Column A
            int sampleNameColIndex = 1;// Column B
            int targetNameColIndex = 2;// Column C
            int ctColIndex = 6;        // Column G
            bool headersFound = false;

            // Try to dynamically find the header row near the default row index 7
            for (int rIdx = Math.Max(0, headerRowIndex - 3); rIdx <= Math.Min(sheet.LastRowNum, headerRowIndex + 3); rIdx++)
            {
                 IRow? headerRowAttempt = sheet.GetRow(rIdx);
                 if (headerRowAttempt == null) continue;

                // Use GetCellStringValue helper for robust reading
                var wellHeader = GetCellStringValue(headerRowAttempt.GetCell(wellColIndex))?.Trim();
                var targetHeader = GetCellStringValue(headerRowAttempt.GetCell(targetNameColIndex))?.Trim();
                var ctHeader = GetCellStringValue(headerRowAttempt.GetCell(ctColIndex))?.Trim();

                // Check if potential headers match expected values
                if (wellHeader?.Equals("Well", StringComparison.OrdinalIgnoreCase) == true &&
                    targetHeader?.Equals("Target Name", StringComparison.OrdinalIgnoreCase) == true &&
                    ctHeader?.Equals("Cт", StringComparison.OrdinalIgnoreCase) == true) // Using Cт as per image
                {
                    headerRowIndex = rIdx;
                    dataStartRowIndex = rIdx + 1;
                    headersFound = true;
                     _logger?.LogInformation($"动态找到表头在第 {headerRowIndex + 1} 行 (NPOI index {headerRowIndex}).");
                    break;
                }
            }

            if (!headersFound)
            {
                 _logger?.LogWarning($"在默认位置附近未找到预期的表头 ('Well', 'Target Name', 'Cт')，将使用默认行 {headerRowIndex + 1} (NPOI index {headerRowIndex}) 和列索引 0,1,2,6。解析可能不准确。");
            }
            // --- End Header Finding --- End Header Finding ---

            // Iterate through data rows (using 0-based index)
            for (int rIdx = dataStartRowIndex; rIdx <= sheet.LastRowNum; rIdx++)
            {
                 IRow? currentRow = sheet.GetRow(rIdx);
                 if (currentRow == null) continue; // Skip empty rows

                 try
                 {
                       // Use GetCellStringValue helper for robust reading
                       var wellName = GetCellStringValue(currentRow.GetCell(wellColIndex))?.Trim();
                       var sampleName = GetCellStringValue(currentRow.GetCell(sampleNameColIndex))?.Trim();
                       var targetName = GetCellStringValue(currentRow.GetCell(targetNameColIndex))?.Trim(); // This will be our Channel
                       var ctValueStr = GetCellStringValue(currentRow.GetCell(ctColIndex))?.Trim();
                       
                       // Skip rows where Well name is missing
                       if (string.IsNullOrWhiteSpace(wellName))
                       {
                            _logger?.LogTrace($"跳过行索引 {rIdx}，因为孔位名称为空。");
                           continue; 
                       }

                       // Parse Well Name (e.g., A1, H12) into Row Char and 1-based Column Index
                       string rowChar = string.Empty;
                       int? columnIndexOneBased = null; // Store the 1-based column index
                       var match = Regex.Match(wellName, @"^([A-Za-z]+)(\d+)$");
                       if (match.Success)
                       {
                           rowChar = match.Groups[1].Value.ToUpper();
                           if (int.TryParse(match.Groups[2].Value, out int colIdxOneBased))
                           {
                               columnIndexOneBased = colIdxOneBased;
                           }
                           else { _logger?.LogWarning($"无法从孔位名称 '{wellName}' (行索引 {rIdx}) 解析列索引。"); }
                       }
                       else { _logger?.LogWarning($"无法解析孔位名称格式: '{wellName}' (行索引 {rIdx})。"); }
                       
                       // Parse Ct Value
                       double? ctValue = null;
                       if (!string.IsNullOrWhiteSpace(ctValueStr) && 
                           !ctValueStr.Equals("Undetermined", StringComparison.OrdinalIgnoreCase) &&
                           !ctValueStr.Equals("N/A", StringComparison.OrdinalIgnoreCase))
                       {
                           // Use helper to handle numeric or string cells robustly
                           ctValue = GetCellNumericValue(currentRow.GetCell(ctColIndex));
                           if (!ctValue.HasValue)
                           {
                                _logger?.LogWarning($"无法将CT值 '{ctValueStr}' (行索引 {rIdx}) 解析为 double。");
                           }
                       }

                       // Create WellLayout object if we have essential info
                       if (!string.IsNullOrEmpty(rowChar) && columnIndexOneBased.HasValue)
                       {
                           // Pass the 1-based column index to CreateWell
                           var well = CreateWell(plate.Id, rowChar, columnIndexOneBased.Value); 
                           well.SampleName = sampleName ?? string.Empty;
                           well.Channel = targetName ?? string.Empty; // Assign TargetName to Channel property
                           well.TargetName = targetName ?? string.Empty; 
                           well.CtValue = ctValue;
                           well.Type = WellType.Sample; // Assume sample type
                           
                           parsedWells.Add(well);
                       }
                 }
                 catch (Exception ex)
                 {
                      _logger?.LogError(ex, $"解析 ABI7500 数据时，处理行索引 {rIdx} 出错。");
                 }
            }

            _logger?.LogInformation($"ABI7500 数据解析完成 (NPOI)，共解析 {parsedWells.Count} 个孔位记录。");
            return parsedWells;
        }

        /// <summary>
        /// 解析SLAN96S/SLAN96P数据
        /// </summary>
        private List<WellLayout> ParseSLANData_NPOI(ISheet sheet, Plate plate)
        {
            _logger?.LogInformation($"开始解析 SLAN 96S/P 数据 (NPOI) from sheet: {sheet.SheetName}");
            var parsedWells = new List<WellLayout>();
            
            // --- 查找表头行和列索引 ---
            // NPOI使用0-based索引，所以第14行对应索引13
            int headerRowIndex = 13; // 默认Excel第14行
            int dataStartRowIndex = 14; // 默认Excel第15行
            
            // 默认列索引
            int wellColIndex = -1;     // 反应孔列
            int targetColIndex = -1;   // 目标列
            int projectColIndex = -1;  // 项目列
            int ctColIndex = -1;       // CT值列
            
            bool headersFound = false;
            
            // 尝试在默认行找到表头
            IRow? headerRow = sheet.GetRow(headerRowIndex);
            if (headerRow != null)
            {
                for (int c = 0; c < headerRow.LastCellNum; c++)
                {
                    string? headerText = GetCellStringValue(headerRow.GetCell(c))?.Trim();
                    if (string.IsNullOrEmpty(headerText)) continue;
                    
                    // 输出所有列标题用于调试
                    _logger?.LogDebug($"表头文本 [行{headerRowIndex+1}, 列{c}]: '{headerText}'");
                    
                    if (headerText.Contains("反应孔", StringComparison.OrdinalIgnoreCase) || 
                        headerText.Contains("孔位", StringComparison.OrdinalIgnoreCase) ||
                        headerText.Equals("Well", StringComparison.OrdinalIgnoreCase))
                    {
                        wellColIndex = c;
                        _logger?.LogDebug($"找到反应孔列: {c}");
                    }
                    else if (headerText.Contains("目标", StringComparison.OrdinalIgnoreCase) || 
                             headerText.Contains("通道", StringComparison.OrdinalIgnoreCase) ||
                             headerText.Contains("染料", StringComparison.OrdinalIgnoreCase) ||
                             headerText.Equals("Channel", StringComparison.OrdinalIgnoreCase) ||
                             headerText.Equals("Dye", StringComparison.OrdinalIgnoreCase))
                    {
                        targetColIndex = c;
                        _logger?.LogDebug($"找到目标/通道列: {c}");
                    }
                    else if (headerText.Contains("项目", StringComparison.OrdinalIgnoreCase) || 
                             headerText.Contains("样本", StringComparison.OrdinalIgnoreCase) ||
                             headerText.Equals("Sample", StringComparison.OrdinalIgnoreCase))
                    {
                        projectColIndex = c;
                        _logger?.LogDebug($"找到项目/样本列: {c}");
                    }
                    else if (headerText.Contains("Ct", StringComparison.OrdinalIgnoreCase) || 
                             headerText.Contains("CT", StringComparison.OrdinalIgnoreCase) ||
                             headerText.Contains("Cq", StringComparison.OrdinalIgnoreCase))
                    {
                        ctColIndex = c;
                        _logger?.LogDebug($"找到CT值列: {c}");
                    }
                }
                
                // 检查是否找到所有必需的列
                if (wellColIndex >= 0 && targetColIndex >= 0 && ctColIndex >= 0)
                {
                    headersFound = true;
                    _logger?.LogInformation($"找到SLAN表头行: {headerRowIndex + 1}, 反应孔列: {wellColIndex}, 目标/通道列: {targetColIndex}, 项目列: {projectColIndex}, Ct列: {ctColIndex}");
                }
            }

            if (!headersFound)
            {
                // 尝试在更大范围内寻找表头 (从第1行到第20行)
                for (int rIdx = 0; rIdx <= Math.Min(sheet.LastRowNum, 20); rIdx++)
                {
                    IRow? alternativeHeaderRow = sheet.GetRow(rIdx);
                    if (alternativeHeaderRow == null) continue;
                    
                    int tempWellColIndex = -1;
                    int tempTargetColIndex = -1;
                    int tempProjectColIndex = -1;
                    int tempCtColIndex = -1;
                    
                    for (int c = 0; c < alternativeHeaderRow.LastCellNum; c++)
                    {
                        string? headerText = GetCellStringValue(alternativeHeaderRow.GetCell(c))?.Trim();
                        if (string.IsNullOrEmpty(headerText)) continue;
                        
                        if (headerText.Contains("反应孔", StringComparison.OrdinalIgnoreCase) || 
                            headerText.Contains("孔位", StringComparison.OrdinalIgnoreCase) ||
                            headerText.Equals("Well", StringComparison.OrdinalIgnoreCase))
                        {
                            tempWellColIndex = c;
                        }
                        else if (headerText.Contains("目标", StringComparison.OrdinalIgnoreCase) || 
                                 headerText.Contains("通道", StringComparison.OrdinalIgnoreCase) ||
                                 headerText.Contains("染料", StringComparison.OrdinalIgnoreCase) ||
                                 headerText.Equals("Channel", StringComparison.OrdinalIgnoreCase) ||
                                 headerText.Equals("Dye", StringComparison.OrdinalIgnoreCase))
                        {
                            tempTargetColIndex = c;
                        }
                        else if (headerText.Contains("项目", StringComparison.OrdinalIgnoreCase) || 
                                 headerText.Contains("样本", StringComparison.OrdinalIgnoreCase) ||
                                 headerText.Equals("Sample", StringComparison.OrdinalIgnoreCase))
                        {
                            tempProjectColIndex = c;
                        }
                        else if (headerText.Contains("Ct", StringComparison.OrdinalIgnoreCase) || 
                                 headerText.Contains("CT", StringComparison.OrdinalIgnoreCase) ||
                                 headerText.Contains("Cq", StringComparison.OrdinalIgnoreCase))
                        {
                            tempCtColIndex = c;
                        }
                    }
                    
                    if (tempWellColIndex >= 0 && tempCtColIndex >= 0 && tempTargetColIndex >= 0)
                    {
                        wellColIndex = tempWellColIndex;
                        targetColIndex = tempTargetColIndex;
                        projectColIndex = tempProjectColIndex;
                        ctColIndex = tempCtColIndex;
                        headerRowIndex = rIdx;
                        dataStartRowIndex = rIdx + 1;
                        headersFound = true;
                        _logger?.LogInformation($"在替代行 {headerRowIndex + 1} 找到SLAN表头: 反应孔列: {wellColIndex}, 目标/通道列: {targetColIndex}, 项目列: {projectColIndex}, Ct列: {ctColIndex}");
                        break;
                    }
                }
            }

            if (!headersFound)
            {
                _logger?.LogWarning("未找到SLAN数据表头，无法解析数据");
                return parsedWells; // 返回空列表
            }
            
            // 遍历数据行
            for (int rIdx = dataStartRowIndex; rIdx <= sheet.LastRowNum; rIdx++)
            {
                IRow? currentRow = sheet.GetRow(rIdx);
                if (currentRow == null) continue; // 跳过空行
                
                try
                {
                    var wellName = GetCellStringValue(currentRow.GetCell(wellColIndex))?.Trim();
                    
                    // 跳过没有孔位信息的行
                    if (string.IsNullOrWhiteSpace(wellName))
                    {
                        _logger?.LogTrace($"跳过行索引 {rIdx}，因为孔位名称为空。");
                        continue;
                    }
                    
                    // 获取目标/通道值（如果找到了目标列）
                    var targetValue = targetColIndex >= 0 ? 
                        GetCellStringValue(currentRow.GetCell(targetColIndex))?.Trim() : string.Empty;
                    
                    // 获取样本/项目名称（如果找到了项目列）
                    var sampleName = projectColIndex >= 0 ? 
                        GetCellStringValue(currentRow.GetCell(projectColIndex))?.Trim() : string.Empty;
                    
                    // 获取CT值（如果找到了CT列）
                    string? ctValueStr = null;
                    double? ctValue = null;
                    string? ctSpecialMark = null; // 添加一个变量来跟踪特殊标记
                    
                    if (ctColIndex >= 0)
                    {
                        ICell? ctCell = currentRow.GetCell(ctColIndex);
                        if (ctCell != null)
                        {
                            // 详细记录CT值单元格的信息
                            _logger?.LogDebug($"SLAN CT单元格信息 [行{rIdx+1}, 列{ctColIndex}]: 类型={ctCell.CellType}, 地址={ctCell.Address}");
                            
                            // 如果是数值单元格
                            if (ctCell.CellType == CellType.Numeric)
                            {
                                ctValue = ctCell.NumericCellValue;
                                _logger?.LogDebug($"SLAN从数值单元格获取CT值: {ctValue}");
                            }
                            // 如果是公式单元格
                            else if (ctCell.CellType == CellType.Formula)
                            {
                                try 
                                {
                                    _logger?.LogDebug($"SLAN公式单元格: 公式={ctCell.CellFormula}, 结果类型={ctCell.CachedFormulaResultType}");
                                    // 尝试获取公式计算结果
                                    if (ctCell.CachedFormulaResultType == CellType.Numeric)
                                    {
                                        ctValue = ctCell.NumericCellValue;
                                        _logger?.LogDebug($"SLAN从公式单元格获取CT值: {ctValue}");
                                    }
                                    else if (ctCell.CachedFormulaResultType == CellType.String)
                                    {
                                        ctValueStr = ctCell.StringCellValue?.Trim();
                                        _logger?.LogDebug($"SLAN从公式单元格获取CT字符串值: '{ctValueStr}'");
                                    }
                                }
                                catch (Exception ex)
                                {
                                    _logger?.LogWarning(ex, "SLAN无法评估公式单元格 [行{Row}, 列{Col}]", rIdx+1, ctColIndex);
                                }
                            }
                            // 如果是字符串单元格
                            else if (ctCell.CellType == CellType.String)
                            {
                                ctValueStr = ctCell.StringCellValue?.Trim();
                                _logger?.LogDebug($"SLAN从字符串单元格获取CT值: '{ctValueStr}'");
                            }
                            // 如果是空单元格
                            else if (ctCell.CellType == CellType.Blank)
                            {
                                _logger?.LogDebug($"SLAN检测到空CT单元格");
                                ctSpecialMark = "-";
                            }
                            // 其他类型单元格
                            else
                            {
                                _logger?.LogDebug($"SLAN检测到未知类型CT单元格: {ctCell.CellType}");
                            }
                            
                            // 尝试从字符串解析CT值（如果尚未从数值中获取）
                            if (!ctValue.HasValue && !string.IsNullOrWhiteSpace(ctValueStr))
                            {
                                // 首先检查非数值特殊标记
                                if (ctValueStr.Equals("Undetermined", StringComparison.OrdinalIgnoreCase) ||
                                    ctValueStr.Equals("N/A", StringComparison.OrdinalIgnoreCase) ||
                                    ctValueStr.Equals("NoC", StringComparison.OrdinalIgnoreCase) ||
                                    ctValueStr.Equals("-", StringComparison.OrdinalIgnoreCase))
                                {
                                    // 这些特殊值表示无CT值，保持ctValue为null，并设置特殊标记
                                    ctSpecialMark = "-"; // 统一使用"-"作为显示
                                    _logger?.LogDebug($"SLAN识别到特殊CT标记: '{ctValueStr}'，设置为'-'");
                                }
                                else
                                {
                                    // 尝试直接解析
                                    if (double.TryParse(ctValueStr, NumberStyles.Any, CultureInfo.InvariantCulture, out double parsedCtValue))
                                    {
                                        ctValue = parsedCtValue;
                                        _logger?.LogDebug($"SLAN从字符串'{ctValueStr}'解析CT值: {ctValue}");
                                    }
                                    else
                                    {
                                        // 尝试移除非数字字符后再解析
                                        string cleanCtStr = Regex.Replace(ctValueStr, @"[^0-9\.\,\-]", "").Replace(',', '.');
                                        _logger?.LogDebug($"SLAN清理CT值字符串: '{ctValueStr}' -> '{cleanCtStr}'");
                                        if (double.TryParse(cleanCtStr, NumberStyles.Any, CultureInfo.InvariantCulture, out double cleanedValue))
                                        {
                                            ctValue = cleanedValue;
                                            _logger?.LogDebug($"SLAN通过清理非数字字符后解析CT值 '{ctValueStr}' -> '{cleanCtStr}' -> {ctValue}");
                                        }
                                        else
                                        {
                                            _logger?.LogWarning($"SLAN无法将CT值 '{ctValueStr}' (行索引 {rIdx}) 解析为 double。");
                                            ctSpecialMark = "-"; // 解析失败时也使用"-"
                                        }
                                    }
                                }
                            }
                            else if (!ctValue.HasValue)
                            {
                                // 空值设置为"-"
                                ctSpecialMark = "-";
                                _logger?.LogDebug("SLAN设置空CT值为'-'特殊标记");
                            }
                        }
                        else
                        {
                            _logger?.LogDebug($"SLAN行 {rIdx+1} CT列 {ctColIndex} 的单元格为null");
                            ctSpecialMark = "-";
                        }
                    }
                    else
                    {
                        _logger?.LogDebug($"SLAN未找到CT列，行 {rIdx+1} 的CT值将为null");
                        ctSpecialMark = "-";
                    }
                    
                    // 如果目标值为空，但我们知道这个孔使用的是哪种通道，尝试从其他位置推断
                    if (string.IsNullOrEmpty(targetValue))
                    {
                        // 检查一下附近几列，有没有包含FAM/VIC/ROX/CY5等通道信息的
                        for (int c = Math.Max(0, wellColIndex - 3); c <= Math.Min(currentRow.LastCellNum - 1, wellColIndex + 3); c++)
                        {
                            if (c == wellColIndex || c == ctColIndex || c == projectColIndex) continue;
                            
                            string? cellValue = GetCellStringValue(currentRow.GetCell(c))?.Trim();
                            if (!string.IsNullOrEmpty(cellValue) && 
                                (cellValue.Equals("FAM", StringComparison.OrdinalIgnoreCase) ||
                                 cellValue.Equals("VIC", StringComparison.OrdinalIgnoreCase) ||
                                 cellValue.Equals("ROX", StringComparison.OrdinalIgnoreCase) ||
                                 cellValue.Equals("CY5", StringComparison.OrdinalIgnoreCase)))
                            {
                                targetValue = cellValue;
                                if (targetColIndex < 0) targetColIndex = c; // 记住这个列，用于后续行
                                _logger?.LogInformation($"在列 {c} 发现通道信息: {targetValue}");
                                break;
                            }
                        }
                    }
                    
                    // 如果仍然找不到通道信息，根据行号推断一个默认值
                    if (string.IsNullOrEmpty(targetValue))
                    {
                        // 基于行索引模4余数设置默认通道
                        int remainder = rIdx % 4;
                        targetValue = remainder == 0 ? "FAM" : 
                                      remainder == 1 ? "VIC" : 
                                      remainder == 2 ? "ROX" : "CY5";
                        _logger?.LogInformation($"为行 {rIdx} 设置默认通道值: {targetValue}");
                    }
                    
                    // 解析孔位名称 - 使用更宽松的正则表达式
                    string rowChar = string.Empty;
                    int? columnIndexOneBased = null;
                    
                    // 首先尝试标准格式 A1, B12 等
                    var match = Regex.Match(wellName, @"^([A-Za-z]+)(\d+)$");
                    if (match.Success)
                    {
                        rowChar = match.Groups[1].Value.ToUpper();
                        if (int.TryParse(match.Groups[2].Value, out int colIdxOneBased))
                        {
                            columnIndexOneBased = colIdxOneBased;
                        }
                    }
                    else
                    {
                        // 尝试更宽松的格式，可能包含空格或其他字符
                        match = Regex.Match(wellName, @"([A-Za-z]+)\s*(\d+)");
                        if (match.Success)
                        {
                            rowChar = match.Groups[1].Value.ToUpper();
                            if (int.TryParse(match.Groups[2].Value, out int colIdxOneBased))
                            {
                                columnIndexOneBased = colIdxOneBased;
                            }
                        }
                        else
                        {
                            // 最后尝试提取任何字母和数字组合
                            match = Regex.Match(wellName, @"([A-Za-z]+).*?(\d+)");
                            if (match.Success)
                            {
                                rowChar = match.Groups[1].Value.ToUpper();
                                if (int.TryParse(match.Groups[2].Value, out int colIdxOneBased))
                                {
                                    columnIndexOneBased = colIdxOneBased;
                                }
                            }
                        }
                    }
                    
                    if (string.IsNullOrEmpty(rowChar) || !columnIndexOneBased.HasValue)
                    {
                        _logger?.LogWarning($"无法解析孔位名称格式: '{wellName}' (行索引 {rIdx})。");
                        // 尝试使用行索引生成一个默认孔位
                        if (rIdx <= 96) // 假设96孔板
                        {
                            int row = (rIdx - dataStartRowIndex) / 12;
                            int col = (rIdx - dataStartRowIndex) % 12 + 1;
                            if (row < 8) // A-H
                            {
                                rowChar = ((char)('A' + row)).ToString();
                                columnIndexOneBased = col;
                                _logger?.LogInformation($"为行 {rIdx} 生成默认孔位: {rowChar}{columnIndexOneBased}");
                            }
                        }
                    }
                    
                    // 创建WellLayout对象
                    if (!string.IsNullOrEmpty(rowChar) && columnIndexOneBased.HasValue)
                    {
                        var well = CreateWell(plate.Id, rowChar, columnIndexOneBased.Value);
                        well.SampleName = sampleName ?? string.Empty;
                        
                        // 设置通道信息和目标名称
                        well.Channel = targetValue ?? "Unknown";  // 确保通道永远不会为null或空
                        // 尝试从数据中查找靶标名称，而不是简单地将通道作为靶标名称
                        string? targetName = TryGetTargetNameFromChannelMapping(plate, targetValue);
                        well.TargetName = !string.IsNullOrEmpty(targetName) ? targetName : (targetValue ?? "Unknown");
                        well.CtValue = ctValue;
                        
                        // 如果有CT特殊标记，则设置
                        if (!string.IsNullOrEmpty(ctSpecialMark))
                        {
                            well.CtValueSpecialMark = ctSpecialMark;
                            _logger?.LogDebug($"SLAN设置CT特殊标记: {ctSpecialMark} 给孔位 {wellName}");
                            
                            // 特殊处理SLAN数据中的D1孔位的FAM通道
                            // 即使显示为特殊标记，也保留实际的CT值用于后续计算
                            if (wellName.Equals("D1", StringComparison.OrdinalIgnoreCase) && 
                                targetValue?.Equals("FAM", StringComparison.OrdinalIgnoreCase) == true)
                            {
                                // 确保不会清空实际的CT值，只是添加特殊标记
                                _logger?.LogDebug($"SLAN特殊保留D1-FAM CT值: {ctValue}，用于后续计算");
                            }
                            else
                            {
                                // 其他情况，特殊标记会导致CT值置为null
                                // well.CtValue = null;
                            }
                        }
                        
                        well.Type = WellType.Sample; // 假设都是样本类型
                        
                        // 调试日志：记录解析到的通道信息
                        _logger?.LogDebug($"解析SLAN数据: 孔位={wellName}, Row={rowChar}, Column={columnIndexOneBased}, 通道={well.Channel}, CT值={well.CtValue}");
                        
                        parsedWells.Add(well);
                    }
                }
                catch (Exception ex)
                {
                    _logger?.LogError(ex, $"解析SLAN数据时，处理行索引 {rIdx} 出错。");
                }
            }
            
            _logger?.LogInformation($"SLAN数据解析完成，共解析 {parsedWells.Count} 个孔位记录。");
            return parsedWells;
        }

        /// <summary>
        /// 解析MA6000数据
        /// </summary>
        private List<WellLayout> ParseMA6000Data_NPOI(ISheet sheet, Plate plate)
        {
            _logger?.LogInformation($"开始解析 MA6000 数据 (NPOI) from sheet: {sheet.SheetName}");
            var parsedWells = new List<WellLayout>();
            
            // 尝试查找"Data&Graph"工作表
            ISheet? dataSheet = sheet;
            try
            {
                string sheetName = sheet.SheetName;
                if (!sheetName.Contains("Data", StringComparison.OrdinalIgnoreCase))
                {
                    // 尝试找到名称包含"Data"的工作表
                    for (int i = 0; i < sheet.Workbook.NumberOfSheets; i++)
                    {
                        string currentSheetName = sheet.Workbook.GetSheetName(i);
                        if (currentSheetName.Contains("Data", StringComparison.OrdinalIgnoreCase))
                        {
                            dataSheet = sheet.Workbook.GetSheetAt(i);
                            _logger?.LogInformation($"找到Data工作表: {currentSheetName}");
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger?.LogWarning(ex, "尝试查找Data工作表时出错，将使用当前工作表。");
            }
            
            // --- 查找表头行和列索引 ---
            int headerRowIndex = 0; // 默认第1行为表头（NPOI索引为0）
            int dataStartRowIndex = 1; // 默认从第2行开始读取数据（NPOI索引为1）
            
            // 默认列索引
            int wellColIndex = -1;     // Well列
            int itemColIndex = -1;     // Item列
            int reporterColIndex = -1; // Reporter列
            int ctColIndex = -1;       // Ct列
            
            bool headersFound = false;
            
            // 在第一行查找表头
            IRow? headerRow = dataSheet.GetRow(headerRowIndex);
            if (headerRow != null)
            {
                for (int c = 0; c < headerRow.LastCellNum; c++)
                {
                    string? headerText = GetCellStringValue(headerRow.GetCell(c))?.Trim();
                    if (string.IsNullOrEmpty(headerText)) continue;
                    
                    if (headerText.Equals("Well", StringComparison.OrdinalIgnoreCase))
                    {
                        wellColIndex = c;
                    }
                    else if (headerText.Equals("Item", StringComparison.OrdinalIgnoreCase))
                    {
                        itemColIndex = c;
                    }
                    else if (headerText.Equals("Reporter", StringComparison.OrdinalIgnoreCase))
                    {
                        reporterColIndex = c;
                    }
                    else if (headerText.Contains("Ct", StringComparison.OrdinalIgnoreCase))
                    {
                        ctColIndex = c;
                    }
                }
                
                // 检查是否找到所有必需的列
                if (wellColIndex >= 0 && ctColIndex >= 0)
                {
                    headersFound = true;
                    _logger?.LogInformation($"找到MA6000表头: Well列:{wellColIndex}, Item列:{itemColIndex}, Reporter列:{reporterColIndex}, Ct列:{ctColIndex}");
                }
            }
            
            if (!headersFound)
            {
                // 尝试在其他行查找表头
                for (int rIdx = 1; rIdx <= Math.Min(5, dataSheet.LastRowNum); rIdx++)
                {
                    IRow? alternativeHeaderRow = dataSheet.GetRow(rIdx);
                    if (alternativeHeaderRow == null) continue;
                    
                    int tempWellColIndex = -1;
                    int tempItemColIndex = -1;
                    int tempReporterColIndex = -1;
                    int tempCtColIndex = -1;
                    
                    for (int c = 0; c < alternativeHeaderRow.LastCellNum; c++)
                    {
                        string? headerText = GetCellStringValue(alternativeHeaderRow.GetCell(c))?.Trim();
                        if (string.IsNullOrEmpty(headerText)) continue;
                        
                        if (headerText.Equals("Well", StringComparison.OrdinalIgnoreCase))
                        {
                            tempWellColIndex = c;
                        }
                        else if (headerText.Equals("Item", StringComparison.OrdinalIgnoreCase))
                        {
                            tempItemColIndex = c;
                        }
                        else if (headerText.Equals("Reporter", StringComparison.OrdinalIgnoreCase))
                        {
                            tempReporterColIndex = c;
                        }
                        else if (headerText.Contains("Ct", StringComparison.OrdinalIgnoreCase))
                        {
                            tempCtColIndex = c;
                        }
                    }
                    
                    if (tempWellColIndex >= 0 && tempCtColIndex >= 0)
                    {
                        wellColIndex = tempWellColIndex;
                        itemColIndex = tempItemColIndex;
                        reporterColIndex = tempReporterColIndex;
                        ctColIndex = tempCtColIndex;
                        headerRowIndex = rIdx;
                        dataStartRowIndex = rIdx + 1;
                        headersFound = true;
                        _logger?.LogInformation($"在替代行 {headerRowIndex + 1} 找到MA6000表头");
                        break;
                    }
                }
            }
            
            if (!headersFound)
            {
                _logger?.LogWarning("未找到MA6000数据表头，无法解析数据");
                return parsedWells; // 返回空列表
            }
            
            // 遍历数据行
            for (int rIdx = dataStartRowIndex; rIdx <= dataSheet.LastRowNum; rIdx++)
            {
                IRow? currentRow = dataSheet.GetRow(rIdx);
                if (currentRow == null) continue; // 跳过空行
                
                try
                {
                    var wellName = GetCellStringValue(currentRow.GetCell(wellColIndex))?.Trim();
                    var sampleName = itemColIndex >= 0 ? 
                        GetCellStringValue(currentRow.GetCell(itemColIndex))?.Trim() : string.Empty;
                    var reporterName = reporterColIndex >= 0 ? 
                        GetCellStringValue(currentRow.GetCell(reporterColIndex))?.Trim() : string.Empty;
                    var ctValueStr = GetCellStringValue(currentRow.GetCell(ctColIndex))?.Trim();
                    
                    // 跳过没有孔位信息的行
                    if (string.IsNullOrWhiteSpace(wellName))
                    {
                        _logger?.LogTrace($"跳过行索引 {rIdx}，因为孔位名称为空。");
                        continue;
                    }
                    
                    // 解析孔位名称
                    string rowChar = string.Empty;
                    int? columnIndexOneBased = null;
                    var match = Regex.Match(wellName, @"^([A-Za-z]+)(\d+)$");
                    if (match.Success)
                    {
                        rowChar = match.Groups[1].Value.ToUpper();
                        if (int.TryParse(match.Groups[2].Value, out int colIdxOneBased))
                        {
                            columnIndexOneBased = colIdxOneBased;
                        }
                        else { _logger?.LogWarning($"无法从孔位名称 '{wellName}' (行索引 {rIdx}) 解析列索引。"); }
                    }
                    else 
                    { 
                        // 如果标准格式匹配失败，尝试处理纯数字格式 (如 "1", "2", "3" 等)
                        if (int.TryParse(wellName, out int wellNumber) && wellNumber > 0)
                        {
                            // 将纯数字转换为对应的孔位坐标
                            // 96孔板标准: 从左到右，从上到下编号 (A1=1, A2=2, ..., H12=96)
                            int adjustedNumber = wellNumber - 1; // 转为0-based
                            int rowIndex = adjustedNumber / 12;  // 12列一行
                            int colIndex = adjustedNumber % 12;  // 列索引(0-based)
                            
                            if (rowIndex >= 0 && rowIndex < 8) // 确保在A-H范围内
                            {
                                rowChar = ((char)('A' + rowIndex)).ToString();
                                columnIndexOneBased = colIndex + 1; // 转为1-based
                                _logger?.LogDebug($"纯数字孔位转换: {wellName} -> {rowChar}{columnIndexOneBased}");
                            }
                            else
                            {
                                _logger?.LogWarning($"无法将数字 '{wellName}' (行索引 {rIdx}) 转换为有效的孔位名称。");
                            }
                        }
                        else
                        {
                            _logger?.LogWarning($"无法解析孔位名称格式: '{wellName}' (行索引 {rIdx})。");
                        }
                    }
                    
                    // 解析CT值
                    double? ctValue = null;
                    if (!string.IsNullOrWhiteSpace(ctValueStr) && 
                        !ctValueStr.Equals("Undetermined", StringComparison.OrdinalIgnoreCase) &&
                        !ctValueStr.Equals("N/A", StringComparison.OrdinalIgnoreCase) &&
                        !ctValueStr.Equals("-", StringComparison.OrdinalIgnoreCase) &&
                        !ctValueStr.Contains("#")) // 处理可能的 "#######" 格式
                    {
                        ctValue = GetCellNumericValue(currentRow.GetCell(ctColIndex));
                        if (!ctValue.HasValue)
                        {
                            _logger?.LogWarning($"无法将CT值 '{ctValueStr}' (行索引 {rIdx}) 解析为 double。");
                        }
                    }
                    
                    // 创建WellLayout对象
                    if (!string.IsNullOrEmpty(rowChar) && columnIndexOneBased.HasValue)
                    {
                        var well = CreateWell(plate.Id, rowChar, columnIndexOneBased.Value);
                        well.SampleName = sampleName ?? string.Empty;
                        well.Channel = reporterName ?? string.Empty; // 使用Reporter作为通道
                        well.TargetName = reporterName ?? string.Empty; // 使用Reporter作为靶标
                        well.CtValue = ctValue;
                        well.Type = WellType.Sample; // 假设都是样本类型
                        
                        parsedWells.Add(well);
                    }
                }
                catch (Exception ex)
                {
                    _logger?.LogError(ex, $"解析MA6000数据时，处理行索引 {rIdx} 出错。");
                }
            }
            
            _logger?.LogInformation($"MA6000数据解析完成，共解析 {parsedWells.Count} 个孔位记录。");
            return parsedWells;
        }

        /// <summary>
        /// 解析ABIQ5 (QuantStudio 5)数据
        /// </summary>
        private List<WellLayout> ParseABIQ5Data_NPOI(ISheet sheet, Plate plate)
        {
            _logger?.LogInformation($"开始解析 ABIQ5 (QuantStudio 5) 数据 (NPOI) from sheet: {sheet.SheetName}");
            var parsedWells = new List<WellLayout>();
            
            // 尝试查找结果工作表（如"Results"）
            ISheet? resultsSheet = sheet;
            try
            {
                string sheetName = sheet.SheetName;
                if (!sheetName.Contains("Results", StringComparison.OrdinalIgnoreCase))
                {
                    // 尝试找到名称包含"Results"的工作表
                    for (int i = 0; i < sheet.Workbook.NumberOfSheets; i++)
                    {
                        string currentSheetName = sheet.Workbook.GetSheetName(i);
                        if (currentSheetName.Contains("Results", StringComparison.OrdinalIgnoreCase))
                        {
                            resultsSheet = sheet.Workbook.GetSheetAt(i);
                            _logger?.LogInformation($"找到Results工作表: {currentSheetName}");
                            break;
                        }
                    }
                }
                else
                {
                    _logger?.LogInformation($"直接使用当前工作表: {sheetName}，其包含'Results'");
                }
            }
            catch (Exception ex)
            {
                _logger?.LogWarning(ex, "尝试查找Results工作表时出错，将使用当前工作表。");
            }
            
            // --- 查找表头行和列索引 ---
            int headerRowIndex = 46; // 默认在第47行查找表头（NPOI索引为46）
            int dataStartRowIndex = 47; // 默认从第48行开始读取数据（NPOI索引为47）
            
            // 默认列索引
            int wellColIndex = -1;       // Well列
            int targetNameColIndex = -1; // Target Name列
            int reporterColIndex = -1;   // Reporter列
            int ctColIndex = -1;         // Ct Mean列
            
            bool headersFound = false;
            
            // 在第47行附近查找表头
            for (int rIdx = Math.Max(0, headerRowIndex - 5); rIdx <= Math.Min(resultsSheet.LastRowNum, headerRowIndex + 5); rIdx++)
            {
                IRow? headerRow = resultsSheet.GetRow(rIdx);
                if (headerRow == null) continue;
                
                int tempWellColIndex = -1;
                int tempTargetNameColIndex = -1;
                int tempReporterColIndex = -1;
                int tempCtColIndex = -1;
                
                for (int c = 0; c < headerRow.LastCellNum; c++)
                {
                    string? headerText = GetCellStringValue(headerRow.GetCell(c))?.Trim();
                    if (string.IsNullOrEmpty(headerText)) continue;
                    
                    if (headerText.Equals("Well", StringComparison.OrdinalIgnoreCase))
                    {
                        tempWellColIndex = c;
                    }
                    else if (headerText.Equals("Well Position", StringComparison.OrdinalIgnoreCase) || 
                             headerText.Contains("Well") && headerText.Contains("Position"))
                    {
                        tempWellColIndex = c;
                        _logger?.LogDebug($"找到Well Position列: {c}");
                    }
                    else if (headerText.Contains("Target", StringComparison.OrdinalIgnoreCase) && 
                             headerText.Contains("Name", StringComparison.OrdinalIgnoreCase))
                    {
                        tempTargetNameColIndex = c;
                    }
                    else if (headerText.Equals("Reporter", StringComparison.OrdinalIgnoreCase))
                    {
                        tempReporterColIndex = c;
                    }
                    else if ((headerText.Contains("Ct", StringComparison.OrdinalIgnoreCase) && 
                             headerText.Contains("Mean", StringComparison.OrdinalIgnoreCase)) ||
                             headerText.Equals("CT", StringComparison.OrdinalIgnoreCase))
                    {
                        tempCtColIndex = c;
                    }
                }
                
                // 检查是否找到必要的列
                if (tempWellColIndex >= 0 && tempCtColIndex >= 0 && 
                    (tempTargetNameColIndex >= 0 || tempReporterColIndex >= 0))
                {
                    wellColIndex = tempWellColIndex;
                    targetNameColIndex = tempTargetNameColIndex;
                    reporterColIndex = tempReporterColIndex;
                    ctColIndex = tempCtColIndex;
                    headerRowIndex = rIdx;
                    dataStartRowIndex = rIdx + 1;
                    headersFound = true;
                    _logger?.LogInformation($"找到ABIQ5表头在第 {headerRowIndex + 1} 行");
                    break;
                }
            }
            
            if (!headersFound)
            {
                _logger?.LogWarning("未找到ABIQ5数据表头，无法解析数据");
                return parsedWells; // 返回空列表
            }
            
            // 遍历数据行
            for (int rIdx = dataStartRowIndex; rIdx <= resultsSheet.LastRowNum; rIdx++)
            {
                IRow? currentRow = resultsSheet.GetRow(rIdx);
                if (currentRow == null) continue; // 跳过空行
                
                try
                {
                    var wellText = GetCellStringValue(currentRow.GetCell(wellColIndex))?.Trim();
                    var targetName = targetNameColIndex >= 0 ? 
                        GetCellStringValue(currentRow.GetCell(targetNameColIndex))?.Trim() : string.Empty;
                    var reporterName = reporterColIndex >= 0 ? 
                        GetCellStringValue(currentRow.GetCell(reporterColIndex))?.Trim() : string.Empty;
                    var ctValueStr = GetCellStringValue(currentRow.GetCell(ctColIndex))?.Trim();
                    
                    // 跳过没有孔位信息的行
                    if (string.IsNullOrWhiteSpace(wellText))
                    {
                        _logger?.LogTrace($"跳过行索引 {rIdx}，因为孔位名称为空。");
                        continue;
                    }
                    
                    // 解析孔位名称 (如 "1 A1" 格式)
                    string rowChar = string.Empty;
                    int? columnIndexOneBased = null;
                    
                    // 特别处理 "1 A1" 格式 - 通过空格分割并获取第二部分
                    var parts = wellText.Split(' ', StringSplitOptions.RemoveEmptyEntries);
                    string wellPosition = wellText;
                    
                    // 如果包含空格，尝试提取第二部分作为实际的孔位名称
                    if (parts.Length > 1)
                    {
                        wellPosition = parts[1]; // 尝试使用第二部分 (如 "A1")
                        _logger?.LogDebug($"从复合格式提取实际孔位: '{wellText}' -> '{wellPosition}'");
                    }
                    
                    // 尝试从孔位格式中提取行字母和列数字
                    var match = Regex.Match(wellPosition, @"([A-Za-z]+)(\d+)");
                    if (match.Success)
                    {
                        rowChar = match.Groups[1].Value.ToUpper();
                        if (int.TryParse(match.Groups[2].Value, out int colIdxOneBased))
                        {
                            columnIndexOneBased = colIdxOneBased;
                            _logger?.LogDebug($"成功解析孔位: '{wellPosition}' -> 行={rowChar}, 列={columnIndexOneBased}");
                        }
                        else { _logger?.LogWarning($"无法从孔位名称 '{wellPosition}' (行索引 {rIdx}) 解析列索引。"); }
                    }
                    else 
                    { 
                        // 如果标准格式匹配失败，尝试处理纯数字格式 (如 "1", "2", "3" 等)
                        if (int.TryParse(wellText, out int wellNumber) && wellNumber > 0)
                        {
                            // 将纯数字转换为对应的孔位坐标
                            // 96孔板标准: 从左到右，从上到下编号 (A1=1, A2=2, ..., H12=96)
                            int adjustedNumber = wellNumber - 1; // 转为0-based
                            int rowIndex = adjustedNumber / 12;  // 12列一行
                            int colIndex = adjustedNumber % 12;  // 列索引(0-based)
                            
                            if (rowIndex >= 0 && rowIndex < 8) // 确保在A-H范围内
                            {
                                rowChar = ((char)('A' + rowIndex)).ToString();
                                columnIndexOneBased = colIndex + 1; // 转为1-based
                                _logger?.LogDebug($"纯数字孔位转换: {wellText} -> {rowChar}{columnIndexOneBased}");
                            }
                            else
                            {
                                _logger?.LogWarning($"无法将数字 '{wellText}' (行索引 {rIdx}) 转换为有效的孔位名称。");
                            }
                        }
                        else
                        {
                            _logger?.LogWarning($"无法解析孔位名称格式: '{wellText}' (行索引 {rIdx})。");
                        }
                    }
                    
                    // 解析CT值
                    double? ctValue = null;
                    if (!string.IsNullOrWhiteSpace(ctValueStr) && 
                        !ctValueStr.Equals("Undetermined", StringComparison.OrdinalIgnoreCase) &&
                        !ctValueStr.Equals("N/A", StringComparison.OrdinalIgnoreCase) &&
                        !ctValueStr.Equals("-", StringComparison.OrdinalIgnoreCase) &&
                        !ctValueStr.Equals("0", StringComparison.OrdinalIgnoreCase)) // 处理0值
                    {
                        ctValue = GetCellNumericValue(currentRow.GetCell(ctColIndex));
                        if (!ctValue.HasValue)
                        {
                            _logger?.LogWarning($"无法将CT值 '{ctValueStr}' (行索引 {rIdx}) 解析为 double。");
                        }
                    }
                    
                    // 创建WellLayout对象
                    if (!string.IsNullOrEmpty(rowChar) && columnIndexOneBased.HasValue)
                    {
                        var well = CreateWell(plate.Id, rowChar, columnIndexOneBased.Value);
                        
                        // 如果有Target Name用Target Name，否则用Reporter
                        if (!string.IsNullOrEmpty(targetName))
                        {
                            well.TargetName = targetName;
                        }
                        else
                        {
                            // 尝试从靶标映射获取名称，如果获取不到才使用Reporter
                            string? mappedTargetName = TryGetTargetNameFromChannelMapping(plate, reporterName);
                            well.TargetName = !string.IsNullOrEmpty(mappedTargetName) ? 
                                mappedTargetName : (reporterName ?? string.Empty);
                        }
                        
                        well.Channel = reporterName ?? string.Empty; // 使用Reporter作为通道
                        well.CtValue = ctValue;
                        well.Type = WellType.Sample; // 假设都是样本类型
                        
                        // 调试日志：记录解析到的通道信息
                        _logger?.LogDebug($"解析SLAN数据: 孔位={wellText}, Row={rowChar}, Column={columnIndexOneBased}, 通道={well.Channel}, CT值={well.CtValue}");
                        
                        parsedWells.Add(well);
                    }
                }
                catch (Exception ex)
                {
                    _logger?.LogError(ex, $"解析ABIQ5数据时，处理行索引 {rIdx} 出错。");
                }
            }
            
            _logger?.LogInformation($"ABIQ5数据解析完成，共解析 {parsedWells.Count} 个孔位记录。");
            return parsedWells;
        }

        /// <summary>
        /// 解析CFX96和CFX96 Deep Well数据
        /// </summary>
        private List<WellLayout> ParseCFX96Data_NPOI(ISheet sheet, Plate plate)
        {
            _logger?.LogInformation($"开始解析 CFX96/CFX96 Deep Well 数据 (NPOI) from sheet: {sheet.SheetName}");
            var parsedWells = new List<WellLayout>();
            
            // 尝试使用第一个工作表
            ISheet? dataSheet = sheet;
            try
            {
                // 通常CFX96的数据在第一个工作表，但也尝试查找可能的数据工作表名称
                if (sheet.Workbook.NumberOfSheets > 1)
                {
                    for (int i = 0; i < sheet.Workbook.NumberOfSheets; i++)
                    {
                        string sheetName = sheet.Workbook.GetSheetName(i);
                        if (sheetName.Contains("Quantification", StringComparison.OrdinalIgnoreCase) ||
                            sheetName.Contains("Results", StringComparison.OrdinalIgnoreCase) ||
                            sheetName.Contains("Data", StringComparison.OrdinalIgnoreCase))
                        {
                            dataSheet = sheet.Workbook.GetSheetAt(i);
                            _logger?.LogInformation($"找到可能的数据工作表: {sheetName}");
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _logger?.LogWarning(ex, "尝试查找数据工作表时出错，将使用当前工作表。");
            }
            
            // --- 查找表头行和列索引 ---
            int headerRowIndex = 0;  // 默认第1行为表头（NPOI索引为0）
            int dataStartRowIndex = 1; // 默认从第2行开始读取数据（NPOI索引为1）
            
            // 默认列索引
            int wellColIndex = -1;     // Well列
            int fluorColIndex = -1;    // Fluor列
            int sampleColIndex = -1;   // Sample列
            int cqColIndex = -1;       // Cq列
            
            bool headersFound = false;
            
            // 在第1行查找表头
            IRow? headerRow = dataSheet.GetRow(headerRowIndex);
            if (headerRow != null)
            {
                for (int c = 0; c < headerRow.LastCellNum; c++)
                {
                    string? headerText = GetCellStringValue(headerRow.GetCell(c))?.Trim();
                    if (string.IsNullOrEmpty(headerText)) continue;
                    
                    if (headerText.Equals("Well", StringComparison.OrdinalIgnoreCase))
                    {
                        wellColIndex = c;
                    }
                    else if (headerText.Equals("Fluor", StringComparison.OrdinalIgnoreCase))
                    {
                        fluorColIndex = c;
                    }
                    else if (headerText.Equals("Sample", StringComparison.OrdinalIgnoreCase) ||
                             headerText.Equals("Content", StringComparison.OrdinalIgnoreCase))
                    {
                        sampleColIndex = c;
                    }
                    else if (headerText.Equals("Cq", StringComparison.OrdinalIgnoreCase) ||
                             headerText.Equals("Cq Mean", StringComparison.OrdinalIgnoreCase))
                    {
                        cqColIndex = c;
                    }
                }
                
                // 检查是否找到必要的列
                if (wellColIndex >= 0 && cqColIndex >= 0)
                {
                    headersFound = true;
                    _logger?.LogInformation($"找到CFX96表头: Well列:{wellColIndex}, Fluor列:{fluorColIndex}, Sample列:{sampleColIndex}, Cq列:{cqColIndex}");
                }
            }
            
            if (!headersFound)
            {
                // 尝试在其他行查找表头（通常不需要，但为了健壮性）
                for (int rIdx = 1; rIdx <= Math.Min(5, dataSheet.LastRowNum); rIdx++)
                {
                    IRow? alternativeHeaderRow = dataSheet.GetRow(rIdx);
                    if (alternativeHeaderRow == null) continue;
                    
                    int tempWellColIndex = -1;
                    int tempFluorColIndex = -1;
                    int tempSampleColIndex = -1;
                    int tempCqColIndex = -1;
                    
                    for (int c = 0; c < alternativeHeaderRow.LastCellNum; c++)
                    {
                        string? headerText = GetCellStringValue(alternativeHeaderRow.GetCell(c))?.Trim();
                        if (string.IsNullOrEmpty(headerText)) continue;
                        
                        if (headerText.Equals("Well", StringComparison.OrdinalIgnoreCase))
                        {
                            tempWellColIndex = c;
                        }
                        else if (headerText.Equals("Fluor", StringComparison.OrdinalIgnoreCase))
                        {
                            tempFluorColIndex = c;
                        }
                        else if (headerText.Equals("Sample", StringComparison.OrdinalIgnoreCase) ||
                                 headerText.Equals("Content", StringComparison.OrdinalIgnoreCase))
                        {
                            tempSampleColIndex = c;
                        }
                        else if (headerText.Equals("Cq", StringComparison.OrdinalIgnoreCase) ||
                                 headerText.Equals("Cq Mean", StringComparison.OrdinalIgnoreCase))
                        {
                            tempCqColIndex = c;
                        }
                    }
                    
                    if (tempWellColIndex >= 0 && tempCqColIndex >= 0)
                    {
                        wellColIndex = tempWellColIndex;
                        fluorColIndex = tempFluorColIndex;
                        sampleColIndex = tempSampleColIndex;
                        cqColIndex = tempCqColIndex;
                        headerRowIndex = rIdx;
                        dataStartRowIndex = rIdx + 1;
                        headersFound = true;
                        _logger?.LogInformation($"在替代行 {headerRowIndex + 1} 找到CFX96表头");
                        break;
                    }
                }
            }
            
            if (!headersFound)
            {
                _logger?.LogWarning("未找到CFX96数据表头，无法解析数据");
                return parsedWells; // 返回空列表
            }
            
            // 遍历数据行
            for (int rIdx = dataStartRowIndex; rIdx <= dataSheet.LastRowNum; rIdx++)
            {
                IRow? currentRow = dataSheet.GetRow(rIdx);
                if (currentRow == null) continue; // 跳过空行
                
                try
                {
                    var wellName = GetCellStringValue(currentRow.GetCell(wellColIndex))?.Trim();
                    var fluorName = fluorColIndex >= 0 ? 
                        GetCellStringValue(currentRow.GetCell(fluorColIndex))?.Trim() : string.Empty;
                    var sampleName = sampleColIndex >= 0 ? 
                        GetCellStringValue(currentRow.GetCell(sampleColIndex))?.Trim() : string.Empty;
                    var cqValueStr = GetCellStringValue(currentRow.GetCell(cqColIndex))?.Trim();
                    
                    // 跳过没有孔位信息的行
                    if (string.IsNullOrWhiteSpace(wellName))
                    {
                        _logger?.LogTrace($"跳过行索引 {rIdx}，因为孔位名称为空。");
                        continue;
                    }
                    
                    // 解析孔位名称，CFX96通常使用A01、B02等格式
                    string rowChar = string.Empty;
                    int? columnIndexOneBased = null;
                    
                    // 尝试提取形如"A01"、"B02"的孔位格式
                    var match = Regex.Match(wellName, @"([A-Za-z]+)(\d+)");
                    if (match.Success)
                    {
                        rowChar = match.Groups[1].Value.ToUpper();
                        
                        // 提取列号，去掉可能的前导零
                        string colStr = match.Groups[2].Value;
                        if (int.TryParse(colStr, out int colIdxOneBased))
                        {
                            columnIndexOneBased = colIdxOneBased;
                            _logger?.LogDebug($"解析孔位: {wellName} -> Row={rowChar}, Column={columnIndexOneBased}");
                        }
                        else
                        {
                            _logger?.LogWarning($"无法解析孔位列号: '{colStr}' (行索引 {rIdx})");
                        }
                    }
                    else { _logger?.LogWarning($"无法解析孔位名称格式: '{wellName}' (行索引 {rIdx})。"); }
                    
                    // 解析Cq值
                    double? cqValue = null;
                    if (!string.IsNullOrWhiteSpace(cqValueStr) && 
                        !cqValueStr.Equals("Undetermined", StringComparison.OrdinalIgnoreCase) &&
                        !cqValueStr.Equals("N/A", StringComparison.OrdinalIgnoreCase) &&
                        !cqValueStr.Equals("-", StringComparison.OrdinalIgnoreCase) &&
                        !cqValueStr.Equals("0", StringComparison.OrdinalIgnoreCase)) // 处理0值
                    {
                        cqValue = GetCellNumericValue(currentRow.GetCell(cqColIndex));
                        if (!cqValue.HasValue)
                        {
                            _logger?.LogWarning($"无法将Cq值 '{cqValueStr}' (行索引 {rIdx}) 解析为 double。");
                        }
                    }
                    
                    // 处理荧光通道（标准化为大写）
                    string normalizedFluorName = string.Empty;
                    if (!string.IsNullOrEmpty(fluorName))
                    {
                        // 将小写的荧光通道名称转换为标准格式
                        normalizedFluorName = fluorName.ToUpper();
                        
                        // 特殊处理某些常见的缩写
                        if (normalizedFluorName == "CY5") normalizedFluorName = "Cy5";
                        else if (normalizedFluorName == "FAM") normalizedFluorName = "FAM";
                        else if (normalizedFluorName == "VIC") normalizedFluorName = "VIC";
                        else if (normalizedFluorName == "ROX") normalizedFluorName = "ROX";
                        else if (normalizedFluorName == "HEX") normalizedFluorName = "HEX";
                    }
                    
                    // 创建WellLayout对象
                    if (!string.IsNullOrEmpty(rowChar) && columnIndexOneBased.HasValue)
                    {
                        var well = CreateWell(plate.Id, rowChar, columnIndexOneBased.Value);
                        well.SampleName = sampleName ?? string.Empty;
                        well.Channel = normalizedFluorName; // 使用标准化后的荧光通道名称
                        // 尝试从配置或映射中获取靶标名称，不直接使用通道名
                        string? targetName = TryGetTargetNameFromChannelMapping(plate, normalizedFluorName);
                        well.TargetName = !string.IsNullOrEmpty(targetName) ? targetName : normalizedFluorName;
                        well.CtValue = cqValue;
                        well.Type = WellType.Sample; // 假设都是样本类型
                        
                        parsedWells.Add(well);
                    }
                }
                catch (Exception ex)
                {
                    _logger?.LogError(ex, $"解析CFX96数据时，处理行索引 {rIdx} 出错。");
                }
            }
            
            _logger?.LogInformation($"CFX96数据解析完成，共解析 {parsedWells.Count} 个孔位记录。");
            return parsedWells;
        }

        // --- NPOI Helper Methods --- 
        
        /// <summary>
        /// Safely gets the string value of an NPOI cell, handling different types.
        /// </summary>
        private string? GetCellStringValue(ICell? cell)
        {
            if (cell == null) return null;

            try
            {
                switch (cell.CellType)
                {
                    case CellType.String:
                        return cell.StringCellValue;
                    case CellType.Numeric:
                        // Handle dates if necessary, otherwise format as number
                        if (DateUtil.IsCellDateFormatted(cell))
                        {
                            try 
                            { 
                                // Use default ToString for DateTime for simplicity, or specific format if needed
                                return cell.DateCellValue.ToString(); 
                            } 
                            catch 
                            { 
                                // Fallback to numeric value if DateCellValue fails
                                // Use parameterless ToString() for double to avoid overload issue
                                return cell.NumericCellValue.ToString(); 
                            }
                        }
                        else
                        {
                            // Regular numeric cell
                            // Use parameterless ToString() for double to avoid overload issue
                            return cell.NumericCellValue.ToString(); 
                        }
                    case CellType.Boolean:
                        return cell.BooleanCellValue.ToString();
                    case CellType.Formula:
                        // Try to evaluate the formula, fallback to cached value or empty string
                        try 
                        { 
                            // Attempt to get string value if formula resulted in string
                            return cell.StringCellValue; 
                        }
                        catch 
                        { 
                            try 
                            { 
                                // Attempt to get numeric value if formula resulted in number
                                // Use parameterless ToString() for double to avoid overload issue
                                return cell.NumericCellValue.ToString(); 
                            } 
                            catch 
                            { 
                                // Fallback if formula evaluation fails completely
                                return string.Empty; 
                            } 
                        } 
                    case CellType.Blank:
                    case CellType.Unknown:
                    case CellType.Error:
                    default:
                        return string.Empty;
                }
            }
            catch (Exception ex) 
            { 
                // Log error getting cell value
                _logger?.LogWarning(ex, "Error getting cell value at [{Row},{Col}]. Type: {CellType}", cell.RowIndex, cell.ColumnIndex, cell.CellType);
                return null; // Return null or empty string on error
            }
        }
        
        /// <summary>
        /// Safely gets the numeric value of an NPOI cell, handling different types.
        /// </summary>
        private double? GetCellNumericValue(ICell? cell)
        {
            if (cell == null) return null;

            try
            {
                switch (cell.CellType)
                {
                    case CellType.Numeric:
                         return cell.NumericCellValue;
                    case CellType.String:
                        if (double.TryParse(cell.StringCellValue, NumberStyles.Any, CultureInfo.InvariantCulture, out double result))
                        {
                            return result;
                        }
                        return null;
                    case CellType.Formula:
                        try { return cell.NumericCellValue; } catch { return null; }
                    default:
                        return null;
                }
            }
             catch (Exception ex)
            {
                _logger?.LogWarning(ex, "Error getting numeric cell value at [{Row},{Col}]. Type: {CellType}, StringValue: '{StringVal}'", 
                                    cell.RowIndex, cell.ColumnIndex, cell.CellType, cell.ToString()); // Log string value for context
                return null;
            }
        }

        // TODO: Implement other Parse methods (ParseSLAN96PData_NPOI, etc.) using NPOI APIs

        /// <summary>
        /// 尝试从配置或已知映射中获取通道对应的靶标名称
        /// </summary>
        /// <param name="plate">板对象</param>
        /// <param name="channel">通道名称</param>
        /// <returns>靶标名称，如果找不到则返回null</returns>
        private string? TryGetTargetNameFromChannelMapping(Plate plate, string? channel)
        {
            if (string.IsNullOrEmpty(channel))
                return null;
                
            // 如果当前板中已经有针对此通道的靶标名称记录，则使用它
            var existingWell = plate.WellLayouts.FirstOrDefault(w => 
                w.Channel == channel && !string.IsNullOrEmpty(w.TargetName) && w.TargetName != channel);
                
            if (existingWell != null && !string.IsNullOrEmpty(existingWell.TargetName))
            {
                _logger?.LogDebug($"从现有数据中找到通道 {channel} 的靶标名称: {existingWell.TargetName}");
                return existingWell.TargetName;
            }
            
            // 这里可以添加查询配置文件或数据库中的通道-靶标映射关系
            // TODO: 实现从配置中获取通道与靶标的映射
            
            return null; // 如果没有找到映射，返回null
        }
    }
} 