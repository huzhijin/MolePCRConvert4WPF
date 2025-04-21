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

        public async Task<string> GenerateExcelReportAsync(Plate plate, ReportTemplate template, string outputPath)
        {
            _logger.LogInformation("开始生成 Excel 报告: Plate={PlateName}, Template={TemplateName}", plate.Name, template.Name);

            string reportFileName = $"Report_{plate.Name}_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
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
                    // TODO: Implement the actual report population logic here.
                    // This will involve finding specific worksheets/cells/tables in the template
                    // and filling them with data from the 'plate' object (WellLayouts, Samples, PCRResultEntry etc.)
                    // Example: Find a worksheet and write basic plate info
                    var worksheet = package.Workbook.Worksheets.FirstOrDefault(); // Or find by name
                    if (worksheet != null)
                    {                        
                        worksheet.Cells["B2"].Value = plate.Name; // Example cell
                        worksheet.Cells["B3"].Value = plate.InstrumentType.ToString(); // Example cell
                        worksheet.Cells["B4"].Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"); // Example cell
                        
                        _logger.LogInformation("正在填充报告数据到工作表: {WorksheetName}", worksheet.Name);
                        // Add logic to populate well data, sample data, results etc. into the appropriate cells/tables
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
                // Consider throwing a custom exception or returning an error indicator
                throw; // Re-throw the exception for the caller to handle
            }
        }

        public Task<IEnumerable<ReportTemplate>> GetReportTemplatesAsync()
        {
            _logger.LogInformation("获取报告模板列表 (当前实现为占位符)");
            // TODO: Implement logic to discover/load templates from a specific directory or config file
            // Example: Scan a 'Templates' subfolder
            string templateDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates");
            var templates = new List<ReportTemplate>();
            if (Directory.Exists(templateDir))
            {
                var templateFiles = Directory.GetFiles(templateDir, "*.xlsx");
                foreach (var file in templateFiles)
                {
                    templates.Add(new ReportTemplate 
                    { 
                        Id = Guid.NewGuid(), 
                        Name = Path.GetFileNameWithoutExtension(file), 
                        FilePath = file, 
                        IsExcelTemplate = true 
                    });
                }
            }
            return Task.FromResult<IEnumerable<ReportTemplate>>(templates);
        }

        public async Task<ReportTemplate?> GetReportTemplateAsync(Guid id)
        {
             _logger.LogInformation("获取指定ID的报告模板 (当前实现为占位符): ID={TemplateId}", id);
             var templates = await GetReportTemplatesAsync();
            return templates.FirstOrDefault(t => t.Id == id);
        }

        public Task<ReportTemplate> SaveReportTemplateAsync(ReportTemplate reportTemplate)
        {
             _logger.LogInformation("保存报告模板 (当前实现为占位符): Name={TemplateName}", reportTemplate.Name);
            // TODO: Implement logic if templates need to be managed/saved
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
                // Try opening the file with EPPlus as a basic validation
                using (var package = new ExcelPackage(new FileInfo(templateFilePath)))
                {
                    isValid = package.Workbook.Worksheets.Any(); // Check if there's at least one worksheet
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
    }
} 