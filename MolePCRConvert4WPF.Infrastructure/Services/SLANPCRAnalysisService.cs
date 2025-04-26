using Microsoft.Extensions.Logging;
using MolePCRConvert4WPF.Core.Models;
using MolePCRConvert4WPF.Core.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Globalization;
using System.Text.RegularExpressions;
using NCalc;

namespace MolePCRConvert4WPF.Infrastructure.Services
{
    /// <summary>
    /// SLAN系列仪器的PCR分析服务实现
    /// </summary>
    public class SLANPCRAnalysisService : PCRAnalysisService
    {
        private readonly ILogger<SLANPCRAnalysisService> _logger;

        public SLANPCRAnalysisService(ILogger<SLANPCRAnalysisService> logger) : base(logger)
        {
            _logger = logger;
        }

        /// <summary>
        /// SLAN仪器专用的D1孔位FAM通道浓度计算
        /// </summary>
        /// <param name="position">孔位</param>
        /// <param name="channelCtValues">该孔位所有通道的CT值</param>
        /// <returns>计算的浓度值，如果不符合条件则返回null</returns>
        protected double? CalculateSLAND1FamConcentration(string position, Dictionary<string, double?> channelCtValues)
        {
            // 只处理D1孔位
            if (!position.Equals("D1", StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            _logger?.LogDebug($"SLAN特殊处理: 检查D1孔位FAM和VIC通道条件");

            // 获取FAM和VIC通道的CT值
            double? famCtValue = null;
            double? vicCtValue = null;

            if (channelCtValues.TryGetValue("FAM", out var famValue))
            {
                famCtValue = famValue;
                _logger?.LogDebug($"SLAN-D1-FAM CT值: {famCtValue}");
            }

            if (channelCtValues.TryGetValue("VIC", out var vicValue))
            {
                vicCtValue = vicValue;
                _logger?.LogDebug($"SLAN-D1-VIC CT值: {vicCtValue}");
            }

            // 检查是否满足判断条件：{FAM}<36&&{VIC}<36
            bool famValid = famCtValue.HasValue && famCtValue.Value < 36;
            bool vicValid = vicCtValue.HasValue && vicCtValue.Value < 36;

            if (famValid && vicValid)
            {
                // 满足条件，使用5.703*2^(36-{CT})公式，其中CT为FAM通道的CT值
                double ctValue = famCtValue.Value;
                double result = 5.703 * Math.Pow(2, 36 - ctValue);
                _logger?.LogDebug($"SLAN-D1-FAM满足条件: FAM={famCtValue}<36, VIC={vicCtValue}<36, 计算浓度={result:F4}");
                return result;
            }
            else
            {
                _logger?.LogDebug($"SLAN-D1-FAM不满足条件: FAM={famCtValue}, VIC={vicCtValue}");
                return null;
            }
        }

        /// <summary>
        /// 修改超类的浓度计算方法，针对SLAN仪器的D1孔位FAM通道使用特殊计算规则
        /// </summary>
        protected override double? CalculateMultiChannelConcentration(string position, string currentChannel, double? currentCtValue, 
                                                         string? formula, Dictionary<string, Dictionary<string, double?>> wellCtValues,
                                                         string? ctValueSpecialMark = null)
        {
            // 特殊处理SLAN仪器D1孔位的FAM通道
            if (position.Equals("D1", StringComparison.OrdinalIgnoreCase) && 
                currentChannel.Equals("FAM", StringComparison.OrdinalIgnoreCase))
            {
                // 获取当前孔位所有通道的CT值
                Dictionary<string, double?> channelCtValues = wellCtValues[position];

                // 使用SLAN特定的D1-FAM计算逻辑
                double? concentration = CalculateSLAND1FamConcentration(position, channelCtValues);
                if (concentration.HasValue)
                {
                    return concentration;
                }
            }

            // 其他情况使用超类的通用计算方法
            return base.CalculateMultiChannelConcentration(position, currentChannel, currentCtValue, formula, wellCtValues, ctValueSpecialMark);
        }

        /// <summary>
        /// 修改超类的检测结果评估方法，针对SLAN仪器的D1孔位FAM通道使用特殊判断规则
        /// </summary>
        protected override string EvaluateMultiChannelDetectionResult(string position, string currentChannel, double? currentCtValue, 
                                                         string? formula, Dictionary<string, Dictionary<string, double?>> wellCtValues,
                                                         string? ctValueSpecialMark = null)
        {
            // 特殊处理SLAN仪器D1孔位的FAM通道
            if (position.Equals("D1", StringComparison.OrdinalIgnoreCase) && 
                currentChannel.Equals("FAM", StringComparison.OrdinalIgnoreCase))
            {
                // 获取当前孔位所有通道的CT值
                Dictionary<string, double?> channelCtValues = wellCtValues[position];

                // 获取FAM和VIC通道的CT值
                double? famCtValue = null;
                double? vicCtValue = null;

                if (channelCtValues.TryGetValue("FAM", out var famValue))
                {
                    famCtValue = famValue;
                }

                if (channelCtValues.TryGetValue("VIC", out var vicValue))
                {
                    vicCtValue = vicValue;
                }

                // 检查是否满足判断条件：{FAM}<36&&{VIC}<36
                bool famValid = famCtValue.HasValue && famCtValue.Value < 36;
                bool vicValid = vicCtValue.HasValue && vicCtValue.Value < 36;

                if (famValid && vicValid)
                {
                    _logger?.LogDebug($"SLAN-D1-FAM满足条件: FAM={famCtValue}<36, VIC={vicCtValue}<36, 判定为阳性");
                    return "阳性";
                }
            }

            // 其他情况使用超类的通用判断方法
            return base.EvaluateMultiChannelDetectionResult(position, currentChannel, currentCtValue, formula, wellCtValues, ctValueSpecialMark);
        }
    }
} 