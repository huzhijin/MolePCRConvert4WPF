using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Threading.Tasks;
using MolePCRConvert4WPF.Core.Models;

namespace MolePCRConvert4WPF.App.Services
{
    /// <summary>
    /// 模板设计服务，负责模板的保存、加载和生成报告
    /// </summary>
    public class TemplateDesignService
    {
        private const string TemplateFolder = "CustomTemplates";

        /// <summary>
        /// 加载模板
        /// </summary>
        /// <param name="filePath">模板文件路径</param>
        /// <returns>加载的模板</returns>
        public async Task<ReportCustomTemplate> LoadTemplateAsync(string filePath)
        {
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException("指定的模板文件不存在", filePath);
            }

            using (FileStream stream = new FileStream(filePath, FileMode.Open))
            {
                var serializer = new DataContractJsonSerializer(typeof(ReportCustomTemplate));
                return await Task.Run(() => (ReportCustomTemplate)serializer.ReadObject(stream));
            }
        }

        /// <summary>
        /// 保存模板
        /// </summary>
        /// <param name="filePath">保存路径</param>
        /// <param name="template">要保存的模板</param>
        public async Task SaveTemplateAsync(string filePath, ReportCustomTemplate template)
        {
            // 确保目录存在
            string directory = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            using (FileStream stream = new FileStream(filePath, FileMode.Create))
            {
                var serializer = new DataContractJsonSerializer(typeof(ReportCustomTemplate));
                await Task.Run(() => serializer.WriteObject(stream, template));
            }
        }

        /// <summary>
        /// 获取默认模板目录
        /// </summary>
        /// <returns>默认模板目录路径</returns>
        public string GetDefaultTemplateDirectory()
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string templateDir = Path.Combine(baseDir, TemplateFolder);
            
            if (!Directory.Exists(templateDir))
            {
                Directory.CreateDirectory(templateDir);
            }
            
            return templateDir;
        }

        /// <summary>
        /// 生成报告
        /// </summary>
        /// <param name="template">报告模板</param>
        /// <param name="data">PCR结果数据</param>
        /// <param name="outputPath">输出路径</param>
        /// <returns>成功生成的报告路径</returns>
        public async Task<string> GenerateReportAsync(ReportCustomTemplate template, List<PCRResultEntry> data, string outputPath)
        {
            // 这里需要使用适当的库生成Excel或PDF报告
            // 目前使用Excel的互操作库或OpenXML进行实现
            
            // TODO: 实现报告生成逻辑，使用模板和数据
            
            await Task.Delay(100); // 临时占位
            
            return outputPath;
        }

        /// <summary>
        /// 获取所有可用的自定义模板
        /// </summary>
        /// <returns>模板文件列表</returns>
        public List<string> GetAvailableTemplates()
        {
            string templateDir = GetDefaultTemplateDirectory();
            string[] files = Directory.GetFiles(templateDir, "*.rtp");
            return new List<string>(files);
        }
    }
} 