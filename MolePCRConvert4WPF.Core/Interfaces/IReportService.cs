using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using MolePCRConvert4WPF.Core.Models;

namespace MolePCRConvert4WPF.Core.Interfaces
{
    /// <summary>
    /// 报告服务接口
    /// </summary>
    public interface IReportService
    {
        /// <summary>
        /// 生成Excel报告
        /// </summary>
        /// <param name="plate">板数据</param>
        /// <param name="template">报告模板</param>
        /// <param name="outputPath">输出目录路径</param>
        /// <param name="outputFileName">输出文件名</param>
        /// <param name="isPatientReport">是否为患者报告</param>
        /// <param name="analysisResults">分析服务返回的结果列表</param>
        /// <returns>报告文件路径</returns>
        Task<string> GenerateExcelReportAsync(Plate plate, ReportTemplate template, string outputPath, IEnumerable<AnalysisResultItem> analysisResults, string outputFileName = "", bool isPatientReport = false);
        
        /// <summary>
        /// 生成报告预览
        /// </summary>
        /// <param name="plate">板数据</param>
        /// <param name="template">报告模板</param>
        /// <param name="isPatientReport">是否为患者报告</param>
        /// <param name="analysisResults">分析服务返回的结果列表</param>
        /// <returns>HTML格式的预览内容列表和对应的患者名称列表</returns>
        Task<(List<string> HtmlPreviews, List<string> PatientNames)> GenerateReportPreviewAsync(Plate plate, ReportTemplate template, IEnumerable<AnalysisResultItem> analysisResults, bool isPatientReport = false);
        
        /// <summary>
        /// 生成PDF报告 (如果需要)
        /// </summary>
        /// <param name="plate">板数据</param>
        /// <param name="template">报告模板</param>
        /// <param name="outputPath">输出路径</param>
        /// <returns>报告文件路径</returns>
        // Task<string> GeneratePdfReportAsync(Plate plate, ReportTemplate template, string outputPath);
        
        /// <summary>
        /// 获取所有报告模板
        /// </summary>
        /// <returns>报告模板列表</returns>
        Task<IEnumerable<ReportTemplate>> GetReportTemplatesAsync();
        
        /// <summary>
        /// 获取报告模板
        /// </summary>
        /// <param name="id">报告模板ID</param>
        /// <returns>报告模板</returns>
        Task<ReportTemplate?> GetReportTemplateAsync(Guid id);
        
        /// <summary>
        /// 保存报告模板
        /// </summary>
        /// <param name="reportTemplate">报告模板</param>
        /// <returns>保存后的报告模板</returns>
        Task<ReportTemplate> SaveReportTemplateAsync(ReportTemplate reportTemplate);
        
        /// <summary>
        /// 验证报告模板
        /// </summary>
        /// <param name="templateFilePath">模板文件路径</param>
        /// <returns>是否有效</returns>
        Task<bool> ValidateReportTemplateAsync(string templateFilePath);
    }
} 