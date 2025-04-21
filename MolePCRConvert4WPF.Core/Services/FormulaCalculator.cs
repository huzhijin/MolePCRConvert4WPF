using Microsoft.Extensions.Logging;
using MolePCRConvert4WPF.Core.Models;
using NCalc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace MolePCRConvert4WPF.Core.Services
{
    /// <summary>
    /// PCR分析公式计算器，基于NCalc实现
    /// </summary>
    public class FormulaCalculator
    {
        private readonly Dictionary<string, Dictionary<string, double?>> _wellChannelCtValues;
        private readonly ILogger _logger;
        
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="wells">所有孔位数据</param>
        /// <param name="logger">日志记录器</param>
        public FormulaCalculator(IEnumerable<WellLayout> wells, ILogger logger)
        {
            _logger = logger;
            
            // 初始化所有孔位的通道CT值映射 (wellName -> channel -> ctValue)
            _wellChannelCtValues = wells
                .GroupBy(w => w.WellName)
                .ToDictionary(
                    g => g.Key,
                    g => g.ToDictionary(w => w.Channel ?? "", w => w.CtValue)
                );
        }
        
        /// <summary>
        /// 评估阳性判定规则
        /// </summary>
        /// <param name="wellPosition">孔位</param>
        /// <param name="formula">判定公式</param>
        /// <returns>判定结果（null表示无法判定）</returns>
        public bool? EvaluatePositiveRule(string wellPosition, string formula)
        {
            if (string.IsNullOrEmpty(formula)) return null;
            
            try
            {
                // 预处理公式 - 替换通道引用
                string processedFormula = PreprocessFormula(wellPosition, formula);
                
                // 检查是否有未定义的值
                if (processedFormula.Contains("undefined")) return null;
                
                // 使用NCalc计算表达式
                var expr = new Expression(processedFormula);
                
                // 添加自定义函数
                expr.EvaluateFunction += (name, args) => {
                    if (name.ToLower() == "abs")
                    {
                        args.Result = Math.Abs(Convert.ToDouble(args.Parameters[0].Evaluate()));
                    }
                };
                
                object result = expr.Evaluate();
                
                // 检查结果类型并转换为布尔值
                if (result is bool boolResult)
                {
                    return boolResult;
                }
                else if (result is double doubleResult)
                {
                    // 对于数值结果，返回是否不为0
                    return doubleResult != 0;
                }
                
                return null;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "计算公式错误: {Formula}", formula);
                return null;
            }
        }
        
        /// <summary>
        /// 计算浓度值
        /// </summary>
        /// <param name="wellPosition">孔位</param>
        /// <param name="channel">通道</param>
        /// <param name="formula">浓度计算公式</param>
        /// <returns>计算结果</returns>
        public double? CalculateConcentration(string wellPosition, string channel, string formula)
        {
            if (string.IsNullOrEmpty(formula)) return null;
            
            // 检查CT值是否存在
            if (!_wellChannelCtValues.TryGetValue(wellPosition, out var channelValues) || 
                !channelValues.TryGetValue(channel, out var ctValue) || 
                !ctValue.HasValue)
            {
                return null;
            }
            
            try
            {
                // 替换{CT}为实际CT值
                string processedFormula = formula.Replace("{CT}", ctValue.Value.ToString());
                
                // 替换其他通道引用
                processedFormula = PreprocessFormula(wellPosition, processedFormula);
                
                // 修正幂运算符 - NCalc使用Pow函数而不是^
                processedFormula = ConvertPowerOperator(processedFormula);
                
                // 使用NCalc计算表达式
                var expr = new Expression(processedFormula);
                
                // 添加自定义函数
                expr.EvaluateFunction += (name, args) => {
                    if (name.ToLower() == "pow")
                    {
                        double baseValue = Convert.ToDouble(args.Parameters[0].Evaluate());
                        double exponent = Convert.ToDouble(args.Parameters[1].Evaluate());
                        args.Result = Math.Pow(baseValue, exponent);
                    }
                };
                
                object result = expr.Evaluate();
                return Convert.ToDouble(result);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "计算浓度公式错误: {Formula}", formula);
                return null;
            }
        }
        
        /// <summary>
        /// 预处理公式 - 替换通道引用为实际值
        /// </summary>
        private string PreprocessFormula(string wellPosition, string formula)
        {
            // 替换{通道名}格式引用
            return Regex.Replace(formula, @"\{([A-Za-z0-9]+)\}", match => {
                string channelName = match.Groups[1].Value;
                
                if (_wellChannelCtValues.TryGetValue(wellPosition, out var channelValues) && 
                    channelValues.TryGetValue(channelName, out var ctValue) && 
                    ctValue.HasValue)
                {
                    return ctValue.Value.ToString();
                }
                
                return "undefined";
            });
        }
        
        /// <summary>
        /// 转换幂运算符（^ 到 Pow函数）
        /// </summary>
        private string ConvertPowerOperator(string formula)
        {
            // 将 a^b 转换为 Pow(a,b)
            return Regex.Replace(formula, @"(\d+(\.\d+)?)\s*\^\s*(\d+(\.\d+)?)", "Pow($1,$3)");
        }
    }
} 