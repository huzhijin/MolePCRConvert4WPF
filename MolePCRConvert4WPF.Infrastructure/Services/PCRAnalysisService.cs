using Microsoft.Extensions.Logging;
using MolePCRConvert4WPF.Core.Models;
using MolePCRConvert4WPF.Core.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
// using HLC.Expression; // Remove HLC using
using NCalc; // Add NCalc using
using System.Globalization;
using System.Text.RegularExpressions;

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
                    // 1. 创建已配置患者信息的孔位字典
                    Dictionary<string, (string PatientName, string PatientCaseNumber)> configuredWells = 
                        new Dictionary<string, (string, string)>(StringComparer.OrdinalIgnoreCase);

                    // 收集所有已配置患者信息的孔位
                    foreach (var well in plateData.WellLayouts)
                    {
                        if (!string.IsNullOrEmpty(well.PatientName) && !string.IsNullOrEmpty(well.Position))
                        {
                            configuredWells[well.Position] = (well.PatientName, well.PatientCaseNumber ?? "-");
                            _logger.LogDebug("记录已配置患者的孔位: {Position} -> {Patient}", well.Position, well.PatientName);
                        }
                    }

                    _logger.LogInformation("找到{Count}个已配置患者信息的孔位", configuredWells.Count);

                    // 2. 创建已处理孔位的集合，用于跟踪哪些孔位已有数据
                    HashSet<string> processedWells = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    
                    // 新增：创建孔位-通道-CT值映射表，用于跨通道访问CT值
                    Dictionary<string, Dictionary<string, double?>> wellCtValues = new Dictionary<string, Dictionary<string, double?>>(StringComparer.OrdinalIgnoreCase);
                    
                    // 新增：创建孔位-通道-结果映射表，用于关联通道结果
                    Dictionary<string, Dictionary<string, string>> wellResults = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);

                    // 首先收集所有孔位和通道的CT值
                    foreach (var well in plateData.WellLayouts)
                    {
                        string position = well.Position;
                        string? channel = well.Channel;
                        
                        if (string.IsNullOrEmpty(position) || string.IsNullOrEmpty(channel))
                            continue;
                        
                        // 确保孔位在字典中
                        if (!wellCtValues.ContainsKey(position))
                        {
                            wellCtValues[position] = new Dictionary<string, double?>(StringComparer.OrdinalIgnoreCase);
                            wellResults[position] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                        }
                        
                        // 记录该孔位该通道的CT值
                        wellCtValues[position][channel] = well.CtValue;
                        
                        // 记录此孔位已处理
                        processedWells.Add(position);
                    }
                    
                    int matchedRules = 0;
                    int unmatchedWells = 0;
                    
                    // 现在处理每个孔位的每个通道
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
                            
                            // --- 特殊处理以支持多通道公式 ---
                            string detectionResult = "";
                            double? concentration = null;
                            
                            // 如果公式中包含其他通道CT值的引用（如{FAM}或{VIC}）
                            bool isMultiChannelFormula = rule.PositiveCutoffFormula?.Contains("{") == true && 
                                                         (rule.PositiveCutoffFormula.Contains("{FAM}") || 
                                                          rule.PositiveCutoffFormula.Contains("{VIC}") || 
                                                          rule.PositiveCutoffFormula.Contains("{ROX}") || 
                                                          rule.PositiveCutoffFormula.Contains("{CY5}"));
                            
                            if (isMultiChannelFormula)
                            {
                                // 使用具有多通道感知的评估方法
                                detectionResult = EvaluateMultiChannelDetectionResultInternal(
                                    currentPosition, currentChannel, currentCtValue, rule.PositiveCutoffFormula, wellCtValues, well.CtValueSpecialMark);
                            }
                            else
                            {
                                // 使用常规单通道评估方法
                                detectionResult = EvaluateDetectionResultWithNCalc(currentCtValue, rule.PositiveCutoffFormula, well.CtValueSpecialMark);
                            }
                            
                            // 类似地处理浓度计算
                            bool isMultiChannelConcentration = rule.ConcentrationFormula?.Contains("{") == true && 
                                                               (rule.ConcentrationFormula.Contains("{FAM}") || 
                                                                rule.ConcentrationFormula.Contains("{VIC}") || 
                                                                rule.ConcentrationFormula.Contains("{ROX}") || 
                                                                rule.ConcentrationFormula.Contains("{CY5}"));
                            
                            if (isMultiChannelConcentration)
                            {
                                concentration = CalculateMultiChannelConcentrationInternal(
                                    currentPosition, currentChannel, currentCtValue, rule.ConcentrationFormula, wellCtValues, well.CtValueSpecialMark);
                            }
                            else
                            {
                                concentration = CalculateConcentrationWithNCalc(currentCtValue, rule.ConcentrationFormula, well.CtValueSpecialMark);
                            }
                            
                            // 记录该孔位该通道的结果
                            if (!string.IsNullOrEmpty(detectionResult))
                            {
                                wellResults[currentPosition][currentChannel] = detectionResult;
                            }
                            
                            // --- 创建结果项 ---
                            var resultItem = new AnalysisResultItem
                            {
                                // 优先使用已配置的患者信息
                                PatientName = configuredWells.TryGetValue(currentPosition, out var patientInfo) 
                                    ? patientInfo.PatientName 
                                    : (string.IsNullOrEmpty(well.PatientName) ? "未知患者" : well.PatientName),
                                
                                PatientCaseNumber = configuredWells.TryGetValue(currentPosition, out var caseInfo) 
                                    ? caseInfo.PatientCaseNumber 
                                    : (string.IsNullOrEmpty(well.PatientCaseNumber) ? "-" : well.PatientCaseNumber),
                                
                                WellPosition = currentPosition,
                                Channel = currentChannel,
                                TargetName = rule.TargetName ?? rule.SpeciesName ?? "N/A", // 优先使用TargetName，其次是SpeciesName
                                CtValue = currentCtValue, // 直接使用原始CT值，不做处理
                                CtValueSpecialMark = well.CtValueSpecialMark, // 添加特殊标记
                                Concentration = concentration,
                                // 当有特殊标记时，结果显示为"-"
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
                                // 优先使用已配置的患者信息
                                PatientName = configuredWells.TryGetValue(currentPosition, out var patientInfo) 
                                    ? patientInfo.PatientName 
                                    : (string.IsNullOrEmpty(well.PatientName) ? "未知患者" : well.PatientName),
                                
                                PatientCaseNumber = configuredWells.TryGetValue(currentPosition, out var caseInfo) 
                                    ? caseInfo.PatientCaseNumber 
                                    : (string.IsNullOrEmpty(well.PatientCaseNumber) ? "-" : well.PatientCaseNumber),
                                
                                WellPosition = currentPosition,
                                Channel = currentChannel,
                                TargetName = "未知靶标",
                                CtValue = currentCtValue, // 直接使用原始CT值，不做处理
                                CtValueSpecialMark = well.CtValueSpecialMark, // 添加特殊标记
                                Concentration = null,
                                // 当有特殊标记时，结果显示为"-"
                                DetectionResult = !string.IsNullOrEmpty(well.CtValueSpecialMark) ? "-" : "无法判定"
                            };
                            results.Add(resultItem);
                            
                            // 记录几条未匹配结果用于调试
                            if (unmatchedWells <= 3)
                            {
                                _logger.LogDebug("未找到匹配规则: 孔位={WellPos}, 通道={Channel}", currentPosition, currentChannel);
                            }
                        }
                    }
                    
                    // 后处理：处理通道间的关联关系
                    // 例如，让F2孔位的VIC和FAM共享结果，让CY5和ROX共享结果
                    ProcessChannelRelationships(results, wellResults);
                    
                    // 在分析完成后，为没有数据但已配置患者的孔位添加空结果条目
                    foreach (var wellConfig in configuredWells)
                    {
                        string position = wellConfig.Key;
                        var patientInfo = wellConfig.Value;
                        
                        // 如果该孔位没有被处理过(没有导入数据)
                        if (!processedWells.Contains(position))
                        {
                            _logger.LogDebug("为已配置患者但无数据的孔位添加空结果: {Position} -> {Patient}", 
                                position, patientInfo.PatientName);
                                
                            // 为每个标准通道创建一个结果项
                            string[] standardChannels = new[] { "FAM", "VIC", "ROX", "CY5" };
                            
                            foreach (var channel in standardChannels)
                            {
                                var emptyResult = new AnalysisResultItem
                                {
                                    PatientName = patientInfo.PatientName,     // 显示正确的患者名
                                    PatientCaseNumber = patientInfo.PatientCaseNumber,
                                    WellPosition = position,
                                    Channel = channel,
                                    TargetName = "-",                         // 靶标显示为"-"
                                    CtValue = null,                           // CT值为空，将显示为"-"
                                    CtValueSpecialMark = null,                   // 特殊标记为空
                                    Concentration = null,                     // 浓度为空，将显示为"-"
                                    DetectionResult = "未检出"                 // 检测结果为"未检出"
                                };
                                
                                results.Add(emptyResult);
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

        /// <summary>
        /// 处理通道间的关联关系，使相关通道显示相同的结果
        /// </summary>
        private void ProcessChannelRelationships(List<AnalysisResultItem> results, Dictionary<string, Dictionary<string, string>> wellResults)
        {
            _logger?.LogDebug("开始处理通道间的关联关系");
            
            // 获取所有唯一的孔位
            var positions = results.Select(r => r.WellPosition).Distinct().ToList();
            
            foreach (var position in positions)
            {
                // 如果孔位名称以F2开头，应用特殊的通道关联规则
                if (position?.StartsWith("F2", StringComparison.OrdinalIgnoreCase) == true)
                {
                    _logger?.LogDebug($"处理F2孔位通道关联: {position}");
                    
                    // 查找该孔位下的所有通道结果
                    var positionResults = results.Where(r => r.WellPosition == position).ToList();
                    
                    // 找到FAM和VIC通道的结果项
                    var famResult = positionResults.FirstOrDefault(r => r.Channel?.Equals("FAM", StringComparison.OrdinalIgnoreCase) == true);
                    var vicResult = positionResults.FirstOrDefault(r => r.Channel?.Equals("VIC", StringComparison.OrdinalIgnoreCase) == true);
                    
                    // 如果FAM有结果但VIC没有，或者VIC是"参照"，则将FAM的结果复制到VIC
                    if (famResult != null && vicResult != null && 
                        (!string.IsNullOrEmpty(famResult.DetectionResult) && 
                         (string.IsNullOrEmpty(vicResult.DetectionResult) || 
                          vicResult.TargetName?.Contains("参照") == true)))
                    {
                        _logger?.LogDebug($"将F2 FAM结果 '{famResult.DetectionResult}' 复制到F2 VIC");
                        vicResult.DetectionResult = famResult.DetectionResult;
                    }
                    
                    // 找到ROX和CY5通道的结果项
                    var roxResult = positionResults.FirstOrDefault(r => r.Channel?.Equals("ROX", StringComparison.OrdinalIgnoreCase) == true);
                    var cy5Result = positionResults.FirstOrDefault(r => r.Channel?.Equals("CY5", StringComparison.OrdinalIgnoreCase) == true);
                    
                    // 如果ROX有结果但CY5没有，则将ROX的结果复制到CY5
                    if (roxResult != null && cy5Result != null && 
                        !string.IsNullOrEmpty(roxResult.DetectionResult) && 
                        string.IsNullOrEmpty(cy5Result.DetectionResult))
                    {
                        _logger?.LogDebug($"将F2 ROX结果 '{roxResult.DetectionResult}' 复制到F2 CY5");
                        cy5Result.DetectionResult = roxResult.DetectionResult;
                    }
                }
            }
            
            _logger?.LogDebug("通道间关联关系处理完成");
        }

        // 调用受保护虚拟方法的内部适配器方法
        private string EvaluateMultiChannelDetectionResultInternal(string position, string currentChannel, double? currentCtValue, 
                                                         string? formula, Dictionary<string, Dictionary<string, double?>> wellCtValues,
                                                         string? ctValueSpecialMark = null)
        {
            return EvaluateMultiChannelDetectionResult(position, currentChannel, currentCtValue, formula, wellCtValues, ctValueSpecialMark);
        }

        // 调用受保护虚拟方法的内部适配器方法
        private double? CalculateMultiChannelConcentrationInternal(string position, string currentChannel, double? currentCtValue, 
                                                         string? formula, Dictionary<string, Dictionary<string, double?>> wellCtValues,
                                                         string? ctValueSpecialMark = null)
        {
            return CalculateMultiChannelConcentration(position, currentChannel, currentCtValue, formula, wellCtValues, ctValueSpecialMark);
        }

        /// <summary>
        /// Evaluates the detection result using NCalc.
        /// </summary>
        private string EvaluateDetectionResultWithNCalc(double? ctValue, string? positiveCutoffFormula, string? ctValueSpecialMark = null)
        {
            // 如果有特殊标记，检查判定公式是否直接设置为阳性
            if (!string.IsNullOrEmpty(ctValueSpecialMark))
            {
                // 先检查判定公式是否指定为直接阳性
                if (!string.IsNullOrWhiteSpace(positiveCutoffFormula) &&
                    (positiveCutoffFormula.Equals("POS", StringComparison.OrdinalIgnoreCase) || 
                     positiveCutoffFormula.Equals("POSITIVE", StringComparison.OrdinalIgnoreCase)))
                {
                    _logger?.LogDebug("检测结果评估: CT值有特殊标记 '{CtSpecialMark}'，但判定公式指定为直接阳性，返回'阳性'", ctValueSpecialMark);
                    return "阳性";
                }
                else
                {
                    _logger?.LogDebug("检测结果评估: CT值有特殊标记 '{CtSpecialMark}'，返回'-'", ctValueSpecialMark);
                    return "-";
                }
            }
            
            if (string.IsNullOrWhiteSpace(positiveCutoffFormula))
            {
                _logger?.LogDebug("检测结果评估: 未提供判定公式");
                return "无判定规则";
            }
            
            // 记录原始公式，便于调试
            string originalFormula = positiveCutoffFormula;
            
            // 处理特殊情况
            if (positiveCutoffFormula.Equals("#NaN#", StringComparison.OrdinalIgnoreCase))
            {
                _logger?.LogDebug("检测结果评估: 公式是特殊标记 #NaN#，返回'-'");
                return "-";
            }
            
            if (!ctValue.HasValue) // 无CT值
            {
                _logger?.LogDebug("检测结果评估: CT值为null，返回'未检出'");
                return "未检出"; 
            }
            
            if (ctValue.Value <= 0) // CT值小于等于0视为无效
            {
                _logger?.LogDebug($"检测结果评估: CT值无效 (CT={ctValue.Value})，返回'无效'");
                return "无效"; 
            }
            
            if (ctValue.Value > 45) // CT值过大视为未检出
            {
                _logger?.LogDebug($"检测结果评估: CT值过大 (CT={ctValue.Value})，返回'未检出'");
                return "未检出"; 
            }

            try
            {   
                // 支持特殊值处理
                if (positiveCutoffFormula.Equals("NA", StringComparison.OrdinalIgnoreCase) || 
                    positiveCutoffFormula.Equals("N/A", StringComparison.OrdinalIgnoreCase))
                {
                    _logger?.LogDebug("检测结果评估: 公式是特殊标记 NA/N/A，返回'未检出'");
                    return "未检出";
                }
                
                if (positiveCutoffFormula.Equals("POS", StringComparison.OrdinalIgnoreCase) || 
                    positiveCutoffFormula.Equals("POSITIVE", StringComparison.OrdinalIgnoreCase))
                {
                    _logger?.LogDebug("检测结果评估: 公式是特殊标记 POS/POSITIVE，直接返回'阳性'");
                    return "阳性";
                }
                
                if (positiveCutoffFormula.Equals("NEG", StringComparison.OrdinalIgnoreCase) || 
                    positiveCutoffFormula.Equals("NEGATIVE", StringComparison.OrdinalIgnoreCase))
                {
                    _logger?.LogDebug("检测结果评估: 公式是特殊标记 NEG/NEGATIVE，直接返回'阴性'");
                    return "阴性";
                }
                
                // 处理公式中的各种括号格式 - 更新以支持{CT}、{FAM}等格式
                string ncalcExpression = positiveCutoffFormula
                    .Replace("[CT]", "CT")
                    .Replace("{CT}", "CT")
                    .Replace("[[CT]]", "CT") // 处理可能的双层括号
                    .Replace("^", "**"); // 将^转换为**作为幂运算符
                
                // 处理通道变量 - 替换 {FAM}, {VIC}, {ROX}, {CY5} 等
                ncalcExpression = ncalcExpression
                    .Replace("{FAM}", "CT_FAM") // 替换为安全的变量名
                    .Replace("{VIC}", "CT_VIC")
                    .Replace("{ROX}", "CT_ROX")
                    .Replace("{CY5}", "CT_CY5");
                    
                // 处理ABS函数，将其转换为Math.Abs
                if (ncalcExpression.Contains("ABS("))
                {
                    ncalcExpression = ncalcExpression.Replace("ABS(", "Abs(");
                }
                
                // 添加对2**表达式的处理，将其转换为Pow(2, x)格式
                if (ncalcExpression.Contains("**"))
                {
                    // 使用正则表达式匹配 2** 后面的表达式
                    string pattern = @"2\s*\*\*\s*\(\s*([^)]+)\s*\)";
                    ncalcExpression = Regex.Replace(
                        ncalcExpression, 
                        pattern, 
                        match => $"Pow(2, {match.Groups[1].Value})",
                        RegexOptions.IgnoreCase
                    );
                }
                
                // 特殊处理，将 "Pow (2, x)" 还原为 "Pow(2, x)"，移除不必要的空格
                ncalcExpression = Regex.Replace(ncalcExpression, @"Pow\s+\(\s*2\s*,", "Pow(2,");
                
                // 确保常见操作符周围有空格，使解析更可靠
                ncalcExpression = ncalcExpression
                    .Replace("<="," <= ")
                    .Replace(">="," >= ")
                    .Replace("!="," != ")
                    .Replace("<"," < ")
                    .Replace(">"," > ")
                    .Replace("="," = ")
                    .Replace("*", " * ")
                    .Replace("+", " + ")
                    .Replace("-", " - ")
                    .Replace("/", " / ")
                    .Replace("(", " ( ")
                    .Replace(")", " ) ");
                
                // 特殊处理，将 "Pow (2, x)" 还原为 "Pow(2, x)"，移除不必要的空格
                ncalcExpression = Regex.Replace(ncalcExpression, @"Pow\s+\(\s*2\s*,", "Pow(2,");
                
                // 创建表达式对象
                Expression expression = new Expression(ncalcExpression);
                expression.Parameters["CT"] = ctValue.Value; 
                
                // 添加通道CT值参数 - 目前我们只知道当前通道的CT值
                // 实际使用中，应当根据当前处理的孔位和通道，获取对应的CT值
                expression.Parameters["CT_FAM"] = ctValue.Value; // 简化示例，实际应从数据中获取
                expression.Parameters["CT_VIC"] = ctValue.Value;
                expression.Parameters["CT_ROX"] = ctValue.Value;
                expression.Parameters["CT_CY5"] = ctValue.Value;
                
                // 添加数学函数支持
                expression.EvaluateFunction += (name, args) =>
                {
                    try
                    {
                        switch (name.ToUpper())
                        {
                            case "ABS":
                                if (args.Parameters.Length < 1)
                                {
                                    _logger?.LogWarning($"ABS函数缺少参数");
                                    break;
                                }
                                var absArg = Convert.ToDouble(args.Parameters[0].Evaluate());
                                args.Result = Math.Abs(absArg);
                                _logger?.LogDebug($"计算ABS({absArg}) = {args.Result}");
                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger?.LogError(ex, $"自定义函数 {name} 计算出错");
                        args.Result = double.NaN;
                    }
                };
                
                // 记录处理后的公式和变量，便于调试
                _logger?.LogDebug("检测结果评估: 原始公式={Original}, 处理后={Processed}, CT值={CtValue}", 
                    originalFormula, ncalcExpression, ctValue.Value);
                
                object evaluationResult;
                try
                {
                    evaluationResult = expression.Evaluate();
                    _logger?.LogDebug($"公式计算完成，原始结果: {evaluationResult} (类型: {evaluationResult?.GetType().Name ?? "null"})");
                }
                catch (Exception innerEx)
                {
                    _logger?.LogWarning(innerEx, "公式解析失败: {Formula}", ncalcExpression);
                    return "公式错误";
                }
                
                // 根据结果类型进行判断
                if (evaluationResult is bool boolResult)
                {
                    _logger?.LogDebug($"布尔结果直接判断: {boolResult} -> {(boolResult ? "阳性" : "阴性")}");
                    return boolResult ? "阳性" : "阴性";
                }
                else if (evaluationResult is int || evaluationResult is double || evaluationResult is decimal)
                {
                    // 数值结果处理 - 正数为阳性，非正数为阴性
                    if (double.TryParse(evaluationResult.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out double numericResult))
                    {
                        // 处理特殊值
                        if (double.IsNaN(numericResult))
                        {
                            _logger?.LogWarning("数值结果为NaN，无法判定");
                            return "无法判定";
                        }
                        
                        if (double.IsInfinity(numericResult))
                        {
                            _logger?.LogWarning("数值结果为无穷大，无法判定");
                            return "无法判定";
                        }
                        
                        bool isPositive = numericResult > 0;
                        // 记录数值结果判断过程
                        _logger?.LogDebug("数值结果判断: 值={Value}, 判定为={Result}", numericResult, isPositive ? "阳性" : "阴性");
                        return isPositive ? "阳性" : "阴性";
                    }
                }
                
                _logger?.LogWarning("公式未返回布尔或数值结果: 公式=\"{Formula}\", 结果类型={Type}, 值={Value}", 
                    ncalcExpression, evaluationResult?.GetType().Name ?? "null", evaluationResult);
                return "无法判定";
            }
            catch (Exception ex)
            { 
                _logger?.LogError(ex, "公式计算出错: {Formula} with CT={CtValue}", positiveCutoffFormula, ctValue);
                return "公式错误";
            }
        }

        /// <summary>
        /// Calculates concentration using NCalc.
        /// </summary>
        private double? CalculateConcentrationWithNCalc(double? ctValue, string? concentrationFormula, string? ctValueSpecialMark = null)
        {
            // 如果有特殊标记，检查浓度公式是否是一个固定数值
            if (!string.IsNullOrEmpty(ctValueSpecialMark))
            {
                // 如果浓度公式是一个固定数值，直接返回该数值
                if (!string.IsNullOrWhiteSpace(concentrationFormula) && 
                    double.TryParse(concentrationFormula, out double directValue))
                {
                    _logger?.LogDebug("浓度计算: CT值有特殊标记 '{CtSpecialMark}'，但浓度公式是固定值 {Value}，直接返回该值", 
                        ctValueSpecialMark, directValue);
                    return directValue;
                }
                
                _logger?.LogDebug("浓度计算: CT值有特殊标记 '{CtSpecialMark}'，返回null", ctValueSpecialMark);
                return null;
            }
            
            // 检查输入有效性
            if (string.IsNullOrWhiteSpace(concentrationFormula))
            { 
                _logger?.LogDebug("浓度计算: 未提供公式，返回null");
                return null; // 无公式
            }
            
            // 记录原始公式便于调试
            string originalFormula = concentrationFormula;
            
            // 处理特殊标记
            if (concentrationFormula.Equals("#NaN#", StringComparison.OrdinalIgnoreCase))
            {
                _logger?.LogDebug("浓度计算: 公式是特殊标记 #NaN#，返回null");
                return null;
            }
            
            if (!ctValue.HasValue)
            { 
                _logger?.LogDebug("浓度计算: CT值为null，返回null");
                return null; // CT值无效
            }
            
            if (ctValue.Value <= 0 || ctValue.Value > 45)
            { 
                _logger?.LogDebug($"浓度计算: CT值超出有效范围 (CT={ctValue.Value})，返回null");
                return null; // CT值无效
            }

            try
            {   
                // 处理特殊情况
                if (concentrationFormula.Equals("NA", StringComparison.OrdinalIgnoreCase) || 
                    concentrationFormula.Equals("N/A", StringComparison.OrdinalIgnoreCase))
                {
                    _logger?.LogDebug("浓度计算: 公式是特殊标记 NA/N/A，返回null");
                    return null;
                }
                
                // 如果公式是一个固定数值
                if (double.TryParse(concentrationFormula, out double directValue))
                {
                    _logger?.LogDebug($"浓度计算: 公式是固定数值 {directValue}，直接返回");
                    return directValue;
                }
                
                // 标准化表达式 - 处理各种括号和函数名称
                string ncalcExpression = concentrationFormula
                    .Replace("[CT]", "CT")
                    .Replace("{CT}", "CT")
                    .Replace("[[CT]]", "CT") // 双层括号处理
                    .Replace("^", "**") // 将^替换为**作为幂运算符
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
                
                // 处理2**表达式，将其转换为Pow(2, x)格式
                if (ncalcExpression.Contains("**"))
                {
                    // 使用正则表达式匹配 2** 后面的表达式
                    string pattern = @"2\s*\*\*\s*\(\s*([^)]+)\s*\)";
                    ncalcExpression = Regex.Replace(
                        ncalcExpression, 
                        pattern, 
                        match => $"Pow(2, {match.Groups[1].Value})",
                        RegexOptions.IgnoreCase
                    );
                }
                
                // 处理通道变量 - 替换 {FAM}, {VIC}, {ROX}, {CY5} 等
                ncalcExpression = ncalcExpression
                    .Replace("{FAM}", "CT_FAM") // 替换为安全的变量名
                    .Replace("{VIC}", "CT_VIC")
                    .Replace("{ROX}", "CT_ROX")
                    .Replace("{CY5}", "CT_CY5");
                    
                // 在运算符周围添加空格
                ncalcExpression = ncalcExpression
                    .Replace("<="," <= ")
                    .Replace(">="," >= ")
                    .Replace("!="," != ")
                    .Replace("<"," < ")
                    .Replace(">"," > ")
                    .Replace("="," = ")
                    .Replace("*", " * ")
                    .Replace("+", " + ")
                    .Replace("-", " - ")
                    .Replace("/", " / ")
                    .Replace("(", " ( ")
                    .Replace(")", " ) ");
                
                // 特殊处理，将 "Pow (2, x)" 还原为 "Pow(2, x)"，移除不必要的空格
                ncalcExpression = Regex.Replace(ncalcExpression, @"Pow\s+\(\s*2\s*,", "Pow(2,");
                
                // 创建表达式对象
                Expression expression = new Expression(ncalcExpression);
                
                // 添加所有需要的参数和常量
                expression.Parameters["CT"] = ctValue.Value;
                expression.Parameters["E"] = Math.E;
                expression.Parameters["PI"] = Math.PI;
                
                // 添加通道CT值参数 - 目前我们只知道当前通道的CT值
                expression.Parameters["CT_FAM"] = ctValue.Value; // 简化示例，实际应从数据中获取
                expression.Parameters["CT_VIC"] = ctValue.Value;
                expression.Parameters["CT_ROX"] = ctValue.Value;
                expression.Parameters["CT_CY5"] = ctValue.Value;
                
                // 记录处理后的表达式和CT值，便于调试
                _logger?.LogDebug("浓度计算: 原始公式={Original}, 处理后={Processed}, CT值={CtValue}", 
                    originalFormula, ncalcExpression, ctValue.Value);
                
                // 注册自定义函数
                expression.EvaluateFunction += (name, args) =>
                {
                    try
                    {
                        switch (name.ToUpper())
                        {
                            case "EXP":
                                if (args.Parameters.Length < 1)
                                {
                                    _logger?.LogWarning($"EXP函数缺少参数");
                                    break;
                                }
                                var expArg = Convert.ToDouble(args.Parameters[0].Evaluate());
                                args.Result = Math.Exp(expArg);
                                _logger?.LogDebug($"计算EXP({expArg}) = {args.Result}");
                                break;
                            case "LOG":
                            case "LN":
                                if (args.Parameters.Length < 1)
                                {
                                    _logger?.LogWarning($"LOG/LN函数缺少参数");
                                    break;
                                }
                                var logArg = Convert.ToDouble(args.Parameters[0].Evaluate());
                                if (logArg <= 0)
                                {
                                    _logger?.LogWarning($"LOG/LN函数参数必须大于0: {logArg}");
                                    args.Result = double.NaN;
                                }
                                else
                                {
                                    args.Result = Math.Log(logArg);
                                    _logger?.LogDebug($"计算LOG({logArg}) = {args.Result}");
                                }
                                break;
                            case "LOG10":
                                if (args.Parameters.Length < 1)
                                {
                                    _logger?.LogWarning($"LOG10函数缺少参数");
                                    break;
                                }
                                var log10Arg = Convert.ToDouble(args.Parameters[0].Evaluate());
                                if (log10Arg <= 0)
                                {
                                    _logger?.LogWarning($"LOG10函数参数必须大于0: {log10Arg}");
                                    args.Result = double.NaN;
                                }
                                else
                                {
                                    args.Result = Math.Log10(log10Arg);
                                    _logger?.LogDebug($"计算LOG10({log10Arg}) = {args.Result}");
                                }
                                break;
                            case "POW":
                            case "POWER":
                                if (args.Parameters.Length < 2)
                                {
                                    _logger?.LogWarning($"POW/POWER函数缺少参数");
                                    break;
                                }
                                var powBase = Convert.ToDouble(args.Parameters[0].Evaluate());
                                var powExp = Convert.ToDouble(args.Parameters[1].Evaluate());
                                args.Result = Math.Pow(powBase, powExp);
                                _logger?.LogDebug($"计算POW({powBase}, {powExp}) = {args.Result}");
                                break;
                            case "ABS":
                                if (args.Parameters.Length < 1)
                                {
                                    _logger?.LogWarning($"ABS函数缺少参数");
                                    break;
                                }
                                var absArg = Convert.ToDouble(args.Parameters[0].Evaluate());
                                args.Result = Math.Abs(absArg);
                                _logger?.LogDebug($"计算ABS({absArg}) = {args.Result}");
                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger?.LogError(ex, $"自定义函数 {name} 计算出错");
                        args.Result = double.NaN;
                    }
                };
                
                // 计算结果
                object evaluationResult;
                try
                {
                    evaluationResult = expression.Evaluate();
                    _logger?.LogDebug($"公式计算完成，原始结果: {evaluationResult} (类型: {evaluationResult?.GetType().Name ?? "null"})");
                }
                catch (Exception innerEx)
                {
                    _logger?.LogWarning(innerEx, "浓度公式计算失败: {Formula}", ncalcExpression);
                    return null;
                }

                // 将结果转换为可用的数值
                if (evaluationResult != null)
                {
                    // 如果结果是布尔值，转换为0或1
                    if (evaluationResult is bool boolResult)
                    {
                        double numericResult = boolResult ? 1.0 : 0.0;
                        _logger?.LogDebug($"布尔结果转换为数值: {boolResult} -> {numericResult}");
                        return numericResult;
                    }
                    
                    // 尝试将结果转换为数值
                    if (double.TryParse(evaluationResult.ToString(), 
                        System.Globalization.NumberStyles.Any, 
                        System.Globalization.CultureInfo.InvariantCulture, 
                        out double doubleResult))
                    {
                        // 处理特殊值
                        if (double.IsNaN(doubleResult))
                        {
                            _logger?.LogWarning("浓度计算结果为NaN，返回null");
                            return null;
                        }
                        
                        if (double.IsInfinity(doubleResult))
                        {
                            _logger?.LogWarning("浓度计算结果为无穷大，返回null");
                            return null;
                        }
                        
                        // 浓度不应为负数
                        double finalResult = doubleResult < 0 ? 0 : Math.Round(doubleResult, 4);
                        
                        // 记录计算结果，便于调试
                        _logger?.LogDebug("浓度计算结果: {Result} (原始值={Raw})", 
                            finalResult.ToString("E4"), doubleResult);
                        
                        return finalResult;
                    }
                }
                
                _logger?.LogWarning("浓度公式未返回数值结果: 公式=\"{Formula}\", 结果类型={Type}, 值={Value}", 
                    ncalcExpression, evaluationResult?.GetType().Name ?? "null", evaluationResult);
                return null;
            }
            catch (Exception ex)
            { 
                _logger?.LogError(ex, "浓度计算出错: {Formula} with CT={CtValue}", concentrationFormula, ctValue);
                return null;
            }
        }

        /// <summary>
        /// 评估使用多通道数据的检测结果
        /// </summary>
        protected virtual string EvaluateMultiChannelDetectionResult(string position, string currentChannel, double? currentCtValue, 
                                                         string? formula, Dictionary<string, Dictionary<string, double?>> wellCtValues,
                                                         string? ctValueSpecialMark = null)
        {
            // 如果有特殊标记，直接返回"-"
            if (!string.IsNullOrEmpty(ctValueSpecialMark))
            {
                _logger?.LogDebug($"多通道检测结果评估: 孔位={position}, 通道={currentChannel}, CT值有特殊标记 '{ctValueSpecialMark}'，返回'-'");
                return "-";
            }
            
            if (string.IsNullOrWhiteSpace(formula))
            {
                return "无判定规则";
            }

            _logger?.LogDebug($"多通道公式评估: 孔位={position}, 通道={currentChannel}, 公式={formula}");
            
            try
            {
                // 获取当前孔位所有通道的CT值
                Dictionary<string, double?> channelCtValues = wellCtValues[position];
                
                // 使用NCalc公式评估，但带有多通道CT值支持
                string ncalcExpression = formula
                    .Replace("[CT]", "CT")
                    .Replace("{CT}", "CT")
                    .Replace("[[CT]]", "CT") // 处理可能的双层括号
                    .Replace("^", "**POWER**") // 标记为特殊幂运算符
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
                
                // 处理通道变量 - 替换 {FAM}, {VIC}, {ROX}, {CY5} 等
                ncalcExpression = ncalcExpression
                    .Replace("{FAM}", "CT_FAM") // 替换为安全的变量名
                    .Replace("{VIC}", "CT_VIC")
                    .Replace("{ROX}", "CT_ROX")
                    .Replace("{CY5}", "CT_CY5");
                    
                // 在运算符周围添加空格
                ncalcExpression = ncalcExpression
                    .Replace("*", " * ")
                    .Replace("+", " + ")
                    .Replace("-", " - ")
                    .Replace("/", " / ")
                    .Replace("(", " ( ")
                    .Replace(")", " ) ")
                    .Replace("**POWER**", "**"); // 最后再替换回**幂运算符
                
                // 添加对2**表达式的处理，将其转换为Pow(2, x)格式
                if (ncalcExpression.Contains("**"))
                {
                    // 使用正则表达式匹配 2** 后面的表达式
                    string pattern = @"2\s*\*\*\s*\(\s*([^)]+)\s*\)";
                    ncalcExpression = Regex.Replace(
                        ncalcExpression, 
                        pattern, 
                        match => $"Pow(2, {match.Groups[1].Value})",
                        RegexOptions.IgnoreCase
                    );
                }
                
                // 特殊处理，将 "Pow (2, x)" 还原为 "Pow(2, x)"，移除不必要的空格
                ncalcExpression = Regex.Replace(ncalcExpression, @"Pow\s+\(\s*2\s*,", "Pow(2,");
                
                // 创建表达式对象
                Expression expression = new Expression(ncalcExpression);
                
                // 添加当前通道的CT值
                expression.Parameters["CT"] = currentCtValue ?? 0;
                
                // 添加所有通道的CT值
                foreach (var channelPair in channelCtValues)
                {
                    string channel = channelPair.Key;
                    double? ctValue = channelPair.Value;
                    
                    switch (channel.ToUpper())
                    {
                        case "FAM":
                            expression.Parameters["CT_FAM"] = ctValue ?? 0;
                            break;
                        case "VIC":
                            expression.Parameters["CT_VIC"] = ctValue ?? 0;
                            break;
                        case "ROX":
                            expression.Parameters["CT_ROX"] = ctValue ?? 0;
                            break;
                        case "CY5":
                            expression.Parameters["CT_CY5"] = ctValue ?? 0;
                            break;
                    }
                }
                
                // 添加数学函数支持
                expression.EvaluateFunction += (name, args) =>
                {
                    try
                    {
                        switch (name.ToUpper())
                        {
                            case "ABS":
                                if (args.Parameters.Length < 1)
                                {
                                    _logger?.LogWarning($"ABS函数缺少参数");
                                    break;
                                }
                                var absArg = Convert.ToDouble(args.Parameters[0].Evaluate());
                                args.Result = Math.Abs(absArg);
                                _logger?.LogDebug($"计算ABS({absArg}) = {args.Result}");
                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger?.LogError(ex, $"自定义函数 {name} 计算出错");
                        args.Result = double.NaN;
                    }
                };
                
                // 记录参数值，便于调试
                _logger?.LogDebug($"多通道公式参数: CT={currentCtValue}, CT_FAM={expression.Parameters["CT_FAM"]}, " + 
                                  $"CT_VIC={expression.Parameters["CT_VIC"]}, CT_ROX={expression.Parameters["CT_ROX"]}, CT_CY5={expression.Parameters["CT_CY5"]}");
                
                // 评估表达式
                object result;
                try
                {
                    result = expression.Evaluate();
                }
                catch (Exception ex)
                {
                    _logger?.LogWarning(ex, $"多通道公式评估失败: {ncalcExpression}");
                    return "公式错误";
                }
                
                // 处理结果
                if (result is bool boolResult)
                {
                    _logger?.LogDebug($"多通道公式返回布尔结果: {boolResult} -> {(boolResult ? "阳性" : "阴性")}");
                    return boolResult ? "阳性" : "阴性";
                }
                else if (result is int || result is double || result is decimal)
                {
                    if (double.TryParse(result.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out double numericResult))
                    {
                        if (double.IsNaN(numericResult) || double.IsInfinity(numericResult))
                        {
                            return "无法判定";
                        }
                        
                        bool isPositive = numericResult > 0;
                        _logger?.LogDebug($"多通道公式返回数值结果: {numericResult} -> {(isPositive ? "阳性" : "阴性")}");
                        return isPositive ? "阳性" : "阴性";
                    }
                }
                
                return "无法判定";
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, $"多通道检测结果评估出错: 孔位={position}, 通道={currentChannel}, 公式={formula}");
                return "公式错误";
            }
        }

        /// <summary>
        /// 计算使用多通道数据的浓度
        /// </summary>
        protected virtual double? CalculateMultiChannelConcentration(string position, string currentChannel, double? currentCtValue, 
                                                         string? formula, Dictionary<string, Dictionary<string, double?>> wellCtValues,
                                                         string? ctValueSpecialMark = null)
        {
            // 特殊处理SLAN仪器D1孔位的FAM通道
            // 当D1孔位FAM通道有CT值，但在UI显示为空时，使用实际CT值进行计算
            if (position.Equals("D1", StringComparison.OrdinalIgnoreCase) && 
                currentChannel.Equals("FAM", StringComparison.OrdinalIgnoreCase))
            {
                // 查找FAM和VIC通道是否满足条件{FAM}<36&&{VIC}<36
                bool famValid = false;
                bool vicValid = false;

                // 检查FAM通道值是否<36
                if (wellCtValues[position].TryGetValue("FAM", out var famCtValue) && 
                    famCtValue.HasValue && famCtValue.Value < 36)
                {
                    famValid = true;
                }

                // 检查VIC通道值是否<36
                if (wellCtValues[position].TryGetValue("VIC", out var vicCtValue) && 
                    vicCtValue.HasValue && vicCtValue.Value < 36)
                {
                    vicValid = true;
                }

                // 如果两个通道都满足条件，并且有计算公式
                if (famValid && vicValid && !string.IsNullOrWhiteSpace(formula))
                {
                    // 使用5.703*2^(36-{CT})公式，其中CT为FAM通道的CT值
                    double ctValue = famCtValue.Value;
                    double result = 5.703 * Math.Pow(2, 36 - ctValue);
                    _logger?.LogDebug($"SLAN特殊处理D1孔位FAM通道: 使用CT值={ctValue}计算浓度={result}");
                    return result;
                }
            }

            // 如果有特殊标记，检查浓度公式是否是一个固定数值
            if (!string.IsNullOrEmpty(ctValueSpecialMark))
            {
                // 如果浓度公式是一个固定数值，直接返回该数值
                if (!string.IsNullOrWhiteSpace(formula) && 
                    double.TryParse(formula, out double directValue))
                {
                    _logger?.LogDebug("多通道浓度计算: 孔位={Position}, 通道={Channel}, CT值有特殊标记 '{CtSpecialMark}'，但浓度公式是固定值 {Value}，直接返回该值", 
                        position, currentChannel, ctValueSpecialMark, directValue);
                    return directValue;
                }
                
                _logger?.LogDebug("多通道浓度计算: 孔位={Position}, 通道={Channel}, CT值有特殊标记 '{CtSpecialMark}'，返回null", 
                    position, currentChannel, ctValueSpecialMark);
                return null;
            }
            
            if (string.IsNullOrWhiteSpace(formula))
            {
                return null;
            }
            
            _logger?.LogDebug($"多通道浓度计算: 孔位={position}, 通道={currentChannel}, 公式={formula}");
            
            try
            {
                // 获取当前孔位所有通道的CT值
                Dictionary<string, double?> channelCtValues = wellCtValues[position];
                
                // 使用NCalc公式评估，但带有多通道CT值支持
                string ncalcExpression = formula
                    .Replace("[CT]", "CT")
                    .Replace("{CT}", "CT")
                    .Replace("[[CT]]", "CT") // 处理可能的双层括号
                    .Replace("^", "##POW##") // 替换为临时标记
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
                
                // 处理通道变量 - 替换 {FAM}, {VIC}, {ROX}, {CY5} 等
                ncalcExpression = ncalcExpression
                    .Replace("{FAM}", "CT_FAM") // 替换为安全的变量名
                    .Replace("{VIC}", "CT_VIC")
                    .Replace("{ROX}", "CT_ROX")
                    .Replace("{CY5}", "CT_CY5");
                    
                // 修正: 避免替换普通的 ** 为 ##POW##，因为这会导致2**被错误替换
                if (ncalcExpression.Contains("**"))
                {
                    // 将 2** 表达式替换为 Pow(2, x) 格式
                    string pattern = @"2\s*\*\*\s*\(\s*([^)]+)\s*\)";
                    ncalcExpression = Regex.Replace(
                        ncalcExpression, 
                        pattern, 
                        match => $"Pow(2, {match.Groups[1].Value})",
                        RegexOptions.IgnoreCase
                    );
                }
                
                // 仅在运算符周围添加空格
                ncalcExpression = ncalcExpression
                    .Replace("<="," <= ")
                    .Replace(">="," >= ")
                    .Replace("!="," != ")
                    .Replace("<"," < ")
                    .Replace(">"," > ")
                    .Replace("="," = ")
                    .Replace("*", " * ")
                    .Replace("+", " + ")
                    .Replace("-", " - ")
                    .Replace("/", " / ")
                    .Replace("(", " ( ")
                    .Replace(")", " ) ");
                    
                // 特殊处理，将 "Pow (2, x)" 还原为 "Pow(2, x)"，移除不必要的空格
                ncalcExpression = Regex.Replace(ncalcExpression, @"Pow\s+\(\s*2\s*,", "Pow(2,");
                
                // 创建表达式对象
                Expression expression = new Expression(ncalcExpression);
                
                // 添加当前通道的CT值
                expression.Parameters["CT"] = currentCtValue ?? 0;
                expression.Parameters["E"] = Math.E;
                expression.Parameters["PI"] = Math.PI;
                
                // 添加所有通道的CT值
                foreach (var channelPair in channelCtValues)
                {
                    string channel = channelPair.Key;
                    double? ctValue = channelPair.Value;
                    
                    switch (channel.ToUpper())
                    {
                        case "FAM":
                            expression.Parameters["CT_FAM"] = ctValue ?? 0;
                            break;
                        case "VIC":
                            expression.Parameters["CT_VIC"] = ctValue ?? 0;
                            break;
                        case "ROX":
                            expression.Parameters["CT_ROX"] = ctValue ?? 0;
                            break;
                        case "CY5":
                            expression.Parameters["CT_CY5"] = ctValue ?? 0;
                            break;
                    }
                }
                
                // 添加数学函数支持
                expression.EvaluateFunction += (name, args) =>
                {
                    try
                    {
                        switch (name.ToUpper())
                        {
                            case "EXP":
                                if (args.Parameters.Length < 1) break;
                                var expArg = Convert.ToDouble(args.Parameters[0].Evaluate());
                                args.Result = Math.Exp(expArg);
                                break;
                            case "LOG":
                            case "LN":
                                if (args.Parameters.Length < 1) break;
                                var logArg = Convert.ToDouble(args.Parameters[0].Evaluate());
                                if (logArg <= 0) { args.Result = double.NaN; break; }
                                args.Result = Math.Log(logArg);
                                break;
                            case "LOG10":
                                if (args.Parameters.Length < 1) break;
                                var log10Arg = Convert.ToDouble(args.Parameters[0].Evaluate());
                                if (log10Arg <= 0) { args.Result = double.NaN; break; }
                                args.Result = Math.Log10(log10Arg);
                                break;
                            case "POW":
                            case "POWER":
                                if (args.Parameters.Length < 2) break;
                                var powBase = Convert.ToDouble(args.Parameters[0].Evaluate());
                                var powExp = Convert.ToDouble(args.Parameters[1].Evaluate());
                                args.Result = Math.Pow(powBase, powExp);
                                break;
                            case "ABS":
                                if (args.Parameters.Length < 1) break;
                                var absArg = Convert.ToDouble(args.Parameters[0].Evaluate());
                                args.Result = Math.Abs(absArg);
                                break;
                        }
                    }
                    catch (Exception)
                    {
                        args.Result = double.NaN;
                    }
                };
                
                // 记录参数值，便于调试
                _logger?.LogDebug($"多通道浓度公式参数: CT={currentCtValue}, CT_FAM={expression.Parameters["CT_FAM"]}, " + 
                                  $"CT_VIC={expression.Parameters["CT_VIC"]}, CT_ROX={expression.Parameters["CT_ROX"]}, CT_CY5={expression.Parameters["CT_CY5"]}");
                
                // 评估表达式
                object result;
                try
                {
                    result = expression.Evaluate();
                }
                catch (Exception ex)
                {
                    _logger?.LogWarning(ex, $"多通道浓度公式评估失败: {ncalcExpression}");
                    return null;
                }
                
                // 处理结果
                if (result is bool) // 布尔结果转为0或1
                {
                    return (bool)result ? 1.0 : 0.0;
                }
                else if (result is int || result is double || result is decimal)
                {
                    if (double.TryParse(result.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out double numericResult))
                    {
                        if (double.IsNaN(numericResult) || double.IsInfinity(numericResult))
                        {
                            return null;
                        }
                        
                        // 浓度不应为负数
                        return numericResult < 0 ? 0 : Math.Round(numericResult, 4);
                    }
                }
                
                return null;
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, $"多通道浓度计算出错: 孔位={position}, 通道={currentChannel}, 公式={formula}");
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