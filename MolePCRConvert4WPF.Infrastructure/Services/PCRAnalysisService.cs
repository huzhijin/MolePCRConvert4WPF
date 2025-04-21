using Microsoft.Extensions.Logging;
using MolePCRConvert4WPF.Core.Models;
using MolePCRConvert4WPF.Core.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
// using HLC.Expression; // Remove HLC using
using NCalc; // Add NCalc using

namespace MolePCRConvert4WPF.Infrastructure.Services // Assuming Infrastructure layer for implementation
{
    /// <summary>
    /// Service responsible for performing PCR result analysis calculations.
    /// </summary>
    public class PCRAnalysisService : IPCRAnalysisService
    {
        private readonly ILogger<PCRAnalysisService> _logger;

        // Inject other services if needed (e.g., for complex formula parsing/evaluation)
        public PCRAnalysisService(ILogger<PCRAnalysisService> logger)
        {
            _logger = logger;
        }

        public async Task<List<AnalysisResultItem>> AnalyzeAsync(Plate plateData, AnalysisMethodConfiguration analysisConfig)
        {
            _logger.LogInformation("Starting PCR analysis for Plate ID: {PlateId}", plateData.Id);
            var results = new List<AnalysisResultItem>();

            // Check for nulls before accessing properties
            if (plateData == null)
            {
                 _logger.LogWarning("Analysis cannot proceed: Plate data is null.");
                 return results;
            }
             if (plateData.WellLayouts == null)
            {
                 _logger.LogWarning("Analysis cannot proceed: Plate data WellLayouts is null.");
                 return results;
            }
             if (analysisConfig?.Rules == null)
            {
                _logger.LogWarning("Analysis cannot proceed: Analysis configuration or rules are missing.");
                return results; // Return empty list
            }
            
            _logger.LogInformation("规则列表加载成功，共有{RuleCount}条规则", analysisConfig.Rules.Count);
            
            // 记录几条规则用于调试
            for (int i = 0; i < Math.Min(5, analysisConfig.Rules.Count); i++)
            {
                var rule = analysisConfig.Rules[i];
                _logger.LogDebug("规则样例 #{Index}: 孔位={WellPos}, 通道={Channel}, 靶标={Target}, 判定公式={Judgeformula}", 
                    i+1, rule.WellPosition, rule.Channel, rule.TargetName, rule.PositiveCutoffFormula);
            }

            await Task.Run(() => // Perform potentially long analysis on a background thread
            {
                try
                {
                    int matchedRules = 0;
                    int unmatchedWells = 0;
                    
                    // Directly iterate through each WellLayout, as each represents one channel's data for a well
                    foreach (var well in plateData.WellLayouts)
                    {
                        // Extract data directly from the WellLayout object
                        string currentPosition = well.Position; // Get well position (e.g., "A1")
                        string? currentChannel = well.Channel; // Get the channel for this specific WellLayout entry
                        double? currentCtValue = well.CtValue; // Get the Ct value for this specific WellLayout entry
                        
                        // Validate required data for matching and analysis
                        if (string.IsNullOrEmpty(currentPosition) || string.IsNullOrEmpty(currentChannel))
                        {
                            _logger.LogWarning("Skipping WellLayout entry due to missing Position or Channel. Well ID: {WellId}", well.Id);
                            continue; // Skip this entry if essential info is missing
                        }

                        // Find the matching rule in the configuration
                        var rule = analysisConfig.Rules.FirstOrDefault(r =>
                                    r.WellPositionPatternMatches(currentPosition) && // Use the helper method from AnalysisMethodRule
                                    r.Channel?.Equals(currentChannel, StringComparison.OrdinalIgnoreCase) == true);

                        if (rule != null)
                        {
                            matchedRules++;
                            // --- Perform Calculations based on rule --- 
                            string detectionResult = EvaluateDetectionResultWithNCalc(currentCtValue, rule.PositiveCutoffFormula);
                            double? concentration = CalculateConcentrationWithNCalc(currentCtValue, rule.ConcentrationFormula);

                            var resultItem = new AnalysisResultItem
                            {
                                PatientName = string.IsNullOrEmpty(well.PatientName) ? "未知患者" : well.PatientName,
                                PatientCaseNumber = string.IsNullOrEmpty(well.PatientCaseNumber) ? "-" : well.PatientCaseNumber,
                                WellPosition = currentPosition,
                                Channel = currentChannel,
                                TargetName = rule.TargetName ?? rule.SpeciesName ?? "N/A", // 优先使用TargetName，其次是SpeciesName
                                CtValue = currentCtValue,
                                Concentration = concentration,
                                DetectionResult = detectionResult
                                // IsFirstPatientRow will be set later during processing/sorting in ViewModel
                            };
                            results.Add(resultItem);
                            
                            // 记录几条成功结果用于调试
                            if (matchedRules <= 3)
                            {
                                _logger.LogDebug("成功匹配: 孔位={WellPos}, 通道={Channel}, 靶标={Target}, 结果={Result}, 浓度={Conc}", 
                                    currentPosition, currentChannel, resultItem.TargetName, 
                                    detectionResult, concentration.HasValue ? concentration.Value.ToString("E2") : "-");
                            }
                        }
                        else
                        {
                            unmatchedWells++;
                            // Handle cases where no rule applies to this well/channel combination
                            _logger.LogDebug("No analysis rule found for Well: {WellPos}, Channel: {Channel}", currentPosition, currentChannel);
                            // Create a result item indicating "Not Applicable" or similar
                            var resultItem = new AnalysisResultItem
                            {
                                PatientName = string.IsNullOrEmpty(well.PatientName) ? "未知患者" : well.PatientName,
                                PatientCaseNumber = string.IsNullOrEmpty(well.PatientCaseNumber) ? "-" : well.PatientCaseNumber,
                                WellPosition = currentPosition,
                                Channel = currentChannel,
                                TargetName = "未知靶标",
                                CtValue = currentCtValue,
                                Concentration = null,
                                DetectionResult = "无法判定"
                            };
                            results.Add(resultItem);
                            
                            // 记录几条未匹配结果用于调试
                            if (unmatchedWells <= 3)
                            {
                                _logger.LogDebug("未找到匹配规则: 孔位={WellPos}, 通道={Channel}", currentPosition, currentChannel);
                            }
                        }
                    }
                    
                    _logger.LogInformation("PCR分析完成: 共{TotalCount}条记录, 匹配规则{MatchedCount}条, 未匹配{UnmatchedCount}条", 
                        results.Count, matchedRules, unmatchedWells);
                }
                catch (Exception ex)
                { 
                    _logger.LogError(ex, "Error during PCR analysis calculation loop.");
                    // Consider how to report this error back - maybe throw or return partial results with error info
                }
            });

            _logger.LogInformation("PCR analysis finished for Plate ID: {PlateId}. Generated {ResultCount} results.", plateData.Id, results.Count);
            return results;
        }

        // --- Removed Placeholder/Helper Method GetRawDataForWell --- 
        // The logic is now directly inside AnalyzeAsync based on iterating WellLayouts

        /// <summary>
        /// Evaluates the detection result using NCalc.
        /// </summary>
        private string EvaluateDetectionResultWithNCalc(double? ctValue, string? positiveCutoffFormula)
        {
            if (string.IsNullOrWhiteSpace(positiveCutoffFormula))
            {
                return "无判定规则";
            }
            
            // 记录原始公式，便于调试
            string originalFormula = positiveCutoffFormula;
            
            // 处理特殊情况
            if (positiveCutoffFormula.Equals("#NaN#", StringComparison.OrdinalIgnoreCase))
            {
                return "-";
            }
            
            if (!ctValue.HasValue) // 无CT值
            {
                return "未检出"; 
            }
            
            if (ctValue.Value <= 0 || ctValue.Value > 45) // CT值无效范围
            {
                return "无效"; 
            }

            try
            {   
                // 支持特殊值处理
                if (positiveCutoffFormula.Equals("NA", StringComparison.OrdinalIgnoreCase) || 
                    positiveCutoffFormula.Equals("N/A", StringComparison.OrdinalIgnoreCase))
                {
                    return "未检出";
                }
                
                if (positiveCutoffFormula.Equals("POS", StringComparison.OrdinalIgnoreCase) || 
                    positiveCutoffFormula.Equals("POSITIVE", StringComparison.OrdinalIgnoreCase))
                {
                    return "阳性";
                }
                
                if (positiveCutoffFormula.Equals("NEG", StringComparison.OrdinalIgnoreCase) || 
                    positiveCutoffFormula.Equals("NEGATIVE", StringComparison.OrdinalIgnoreCase))
                {
                    return "阴性";
                }
                
                // 处理公式中的各种括号格式
                string ncalcExpression = positiveCutoffFormula
                    .Replace("[CT]", "CT")
                    .Replace("{CT}", "CT"); 
                
                // 确保常见操作符周围有空格，使解析更可靠
                string processedExpression = ncalcExpression
                    .Replace("<="," <= ")
                    .Replace(">="," >= ")
                    .Replace("!="," != ")
                    .Replace("<"," < ")
                    .Replace(">"," > ")
                    .Replace("="," = ");
                
                // 创建表达式对象
                Expression expression = new Expression(processedExpression);
                expression.Parameters["CT"] = ctValue.Value; 
                
                // 记录处理后的公式和变量，便于调试
                _logger.LogDebug("处理公式: 原始公式={Original}, 处理后={Processed}, CT值={CtValue}", 
                    originalFormula, processedExpression, ctValue.Value);
                
                object evaluationResult;
                try
                {
                    evaluationResult = expression.Evaluate();
                }
                catch (Exception innerEx)
                {
                    _logger.LogWarning(innerEx, "公式解析失败: {Formula}", processedExpression);
                    return "公式错误";
                }
                
                // 根据结果类型进行判断
                if (evaluationResult is bool boolResult)
                {
                    return boolResult ? "阳性" : "阴性";
                }
                else if (evaluationResult is int || evaluationResult is double || evaluationResult is decimal)
                {
                    // 数值结果处理 - 正数为阳性，非正数为阴性
                    double numericResult;
                    if (double.TryParse(evaluationResult.ToString(), out numericResult))
                    {
                        bool isPositive = numericResult > 0;
                        // 记录数值结果判断过程
                        _logger.LogDebug("数值结果判断: 值={Value}, 判定为={Result}", numericResult, isPositive ? "阳性" : "阴性");
                        return isPositive ? "阳性" : "阴性";
                    }
                }
                
                _logger.LogWarning("公式未返回布尔或数值结果: 公式=\"{Formula}\", 结果类型={Type}, 值={Value}", 
                    processedExpression, evaluationResult?.GetType().Name ?? "null", evaluationResult);
                return "无法判定";
            }
            catch (Exception ex)
            { 
                _logger.LogError(ex, "公式计算出错: {Formula} with CT={CtValue}", positiveCutoffFormula, ctValue);
                return "公式错误";
            }
        }

        /// <summary>
        /// Calculates concentration using NCalc.
        /// </summary>
        private double? CalculateConcentrationWithNCalc(double? ctValue, string? concentrationFormula)
        {
            // 检查输入有效性
            if (string.IsNullOrWhiteSpace(concentrationFormula))
            { 
                return null; // 无公式
            }
            
            // 记录原始公式便于调试
            string originalFormula = concentrationFormula;
            
            // 处理特殊标记
            if (concentrationFormula.Equals("#NaN#", StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }
            
            if (!ctValue.HasValue || ctValue.Value <= 0 || ctValue.Value > 45)
            { 
                return null; // CT值无效
            }

            try
            {   
                // 处理特殊情况
                if (concentrationFormula.Equals("NA", StringComparison.OrdinalIgnoreCase) || 
                    concentrationFormula.Equals("N/A", StringComparison.OrdinalIgnoreCase))
                {
                    return null;
                }
                
                // 如果公式是一个固定数值
                if (double.TryParse(concentrationFormula, out double directValue))
                {
                    return directValue;
                }
                
                // 标准化表达式 - 处理各种括号和函数名称
                string ncalcExpression = concentrationFormula
                    .Replace("[CT]", "CT")
                    .Replace("{CT}", "CT")
                    .Replace("EXP(", "Exp(")
                    .Replace("exp(", "Exp(")
                    .Replace("LN(", "Log(")
                    .Replace("ln(", "Log(")
                    .Replace("LOG(", "Log(")
                    .Replace("log(", "Log(")
                    .Replace("LOG10(", "Log10(")
                    .Replace("log10(", "Log10(")
                    .Replace("POW(", "Pow(")
                    .Replace("pow(", "Pow(")
                    .Replace("POWER(", "Pow(")
                    .Replace("power(", "Pow(");
                
                // 创建表达式对象
                Expression expression = new Expression(ncalcExpression);
                
                // 添加所有需要的参数和常量
                expression.Parameters["CT"] = ctValue.Value;
                expression.Parameters["E"] = Math.E;
                expression.Parameters["PI"] = Math.PI;
                
                // 记录处理后的表达式和CT值，便于调试
                _logger.LogDebug("浓度计算: 原始公式={Original}, 处理后={Processed}, CT值={CtValue}", 
                    originalFormula, ncalcExpression, ctValue.Value);
                
                // 注册自定义函数
                expression.EvaluateFunction += (name, args) =>
                {
                    switch (name.ToUpper())
                    {
                        case "EXP":
                            args.Result = Math.Exp(Convert.ToDouble(args.Parameters[0].Evaluate()));
                            break;
                        case "LOG":
                        case "LN":
                            args.Result = Math.Log(Convert.ToDouble(args.Parameters[0].Evaluate()));
                            break;
                        case "LOG10":
                            args.Result = Math.Log10(Convert.ToDouble(args.Parameters[0].Evaluate()));
                            break;
                        case "POW":
                        case "POWER":
                            if (args.Parameters.Length >= 2)
                            {
                                args.Result = Math.Pow(
                                    Convert.ToDouble(args.Parameters[0].Evaluate()),
                                    Convert.ToDouble(args.Parameters[1].Evaluate()));
                            }
                            break;
                    }
                };
                
                // 计算结果
                object evaluationResult;
                try
                {
                    evaluationResult = expression.Evaluate();
                }
                catch (Exception innerEx)
                {
                    _logger.LogWarning(innerEx, "浓度公式计算失败: {Formula}", ncalcExpression);
                    return null;
                }

                // 将结果转换为可用的数值
                if (evaluationResult != null)
                {
                    if (double.TryParse(evaluationResult.ToString(), 
                        System.Globalization.NumberStyles.Any, 
                        System.Globalization.CultureInfo.InvariantCulture, 
                        out double doubleResult))
                    {
                        // 浓度不应为负数
                        double finalResult = doubleResult < 0 ? 0 : Math.Round(doubleResult, 4);
                        
                        // 记录计算结果，便于调试
                        _logger.LogDebug("浓度计算结果: {Result} (原始值={Raw})", 
                            finalResult.ToString("E4"), doubleResult);
                        
                        return finalResult;
                    }
                }
                
                _logger.LogWarning("浓度公式未返回数值结果: 公式=\"{Formula}\", 结果类型={Type}, 值={Value}", 
                    ncalcExpression, evaluationResult?.GetType().Name ?? "null", evaluationResult);
                return null;
            }
            catch (Exception ex)
            { 
                _logger.LogError(ex, "浓度计算出错: {Formula} with CT={CtValue}", concentrationFormula, ctValue);
                return null;
            }
        }

        // Remove or comment out the old placeholder methods
        // private string EvaluateDetectionResult(...) { ... }
        // private double? CalculateConcentration(...) { ... }
        // private bool TryParseCutoff(...) { ... }

         // Helper for WellPositionPatternMatches (assuming it's part of AnalysisMethodRule)
         // If not, implement matching logic here or elsewhere.
         /* 
         private bool WellPositionPatternMatches(string pattern, string position)
         {
             // Implement logic to match patterns like "A1", "B:1-6", "*:12", etc.
             return pattern == position; // Simplistic match
         }
         */
    }
} 