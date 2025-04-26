using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using MolePCRConvert4WPF.Core.Models;

namespace MolePCRConvert4WPF.Core.Interfaces
{
    /// <summary>
    /// 报告模板设计器服务接口
    /// </summary>
    public interface IReportTemplateDesignerService
    {
        /// <summary>
        /// 创建新的报告模板
        /// </summary>
        /// <param name="templateName">模板名称</param>
        /// <returns>新创建的报告模板</returns>
        Task<ReportTemplate> CreateNewTemplateAsync(string templateName);
        
        /// <summary>
        /// 保存报告模板
        /// </summary>
        /// <param name="template">报告模板</param>
        /// <param name="gridData">ReoGrid数据</param>
        /// <returns>保存后的报告模板</returns>
        Task<ReportTemplate> SaveTemplateAsync(ReportTemplate template, byte[] gridData);
        
        /// <summary>
        /// 加载报告模板
        /// </summary>
        /// <param name="template">报告模板</param>
        /// <returns>ReoGrid数据</returns>
        Task<byte[]> LoadTemplateAsync(ReportTemplate template);
        
        /// <summary>
        /// 生成报告
        /// </summary>
        /// <param name="plate">板数据</param>
        /// <param name="template">报告模板</param>
        /// <param name="outputPath">输出目录路径</param>
        /// <param name="outputFileName">输出文件名</param>
        /// <param name="isPatientReport">是否为患者报告</param>
        /// <returns>报告文件路径</returns>
        Task<string> GenerateReportAsync(Plate plate, ReportTemplate template, string outputPath, string outputFileName = "", bool isPatientReport = false);
        
        /// <summary>
        /// 获取所有报告模板
        /// </summary>
        /// <returns>报告模板列表</returns>
        Task<IEnumerable<ReportTemplate>> GetReportTemplatesAsync();
        
        /// <summary>
        /// 删除报告模板
        /// </summary>
        /// <param name="template">报告模板</param>
        /// <returns>是否删除成功</returns>
        Task<bool> DeleteTemplateAsync(ReportTemplate template);
    }
} 