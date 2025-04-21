using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using MolePCRConvert4WPF.Core.Interfaces;
using MolePCRConvert4WPF.Core.Models.PanelRules;

namespace MolePCRConvert4WPF.Core.Services
{
    /// <summary>
    /// Panel规则仓库实现 (基于单个配置文件)
    /// </summary>
    public class PanelRuleRepository : IPanelRuleRepository
    {
        private readonly string _rulesFilePath;
        private PanelRuleConfiguration? _panelRuleConfiguration;
        private readonly JsonSerializerOptions _jsonOptions = new JsonSerializerOptions { WriteIndented = true };

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="rulesFilePath">规则文件路径，默认为应用程序目录下的PanelRules.json</param>
        public PanelRuleRepository(string? rulesFilePath = null)
        {
            _rulesFilePath = rulesFilePath ?? Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "PanelRules.json");
            LoadRules();
        }

        /// <summary>
        /// 加载规则配置
        /// </summary>
        private void LoadRules()
        {
            try
            {
                if (File.Exists(_rulesFilePath))
                {
                    string jsonContent = File.ReadAllText(_rulesFilePath);
                    _panelRuleConfiguration = JsonSerializer.Deserialize<PanelRuleConfiguration>(jsonContent);
                    if (_panelRuleConfiguration == null)
                    {
                        System.Diagnostics.Debug.WriteLine($"警告: 无法反序列化规则文件: {_rulesFilePath}。将使用默认配置。");
                        _panelRuleConfiguration = CreateDefaultConfig();
                    }
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"警告: 规则文件不存在: {_rulesFilePath}。将创建并使用默认配置。");
                    _panelRuleConfiguration = CreateDefaultConfig();
                    SaveChanges(); // 保存默认配置
                }
            }
            catch (Exception ex)
            {
                 System.Diagnostics.Debug.WriteLine($"错误: 加载或创建规则文件时出错: {ex.Message}。将使用默认配置。");
                 _panelRuleConfiguration = CreateDefaultConfig();
            }
        }

        /// <summary>
        /// 获取所有可用的Panel类型名称 (现在只有一个配置，所以Panel名称就是配置名称)
        /// </summary>
        public List<string> GetAvailablePanelTypeNames()
        {
            if (_panelRuleConfiguration != null && !string.IsNullOrEmpty(_panelRuleConfiguration.Name))
            {
                return new List<string> { _panelRuleConfiguration.Name };
            }
            return new List<string>();
        }

        /// <summary>
        /// 获取指定Panel名称的规则配置
        /// </summary>
        public PanelRuleConfiguration? GetPanelRuleConfiguration(string panelName)
        {
            // Since we load only one configuration, check if the name matches
            if (_panelRuleConfiguration != null && _panelRuleConfiguration.Name.Equals(panelName, StringComparison.OrdinalIgnoreCase))
            {
                return _panelRuleConfiguration;
            }
            return null;
        }

        /// <summary>
        /// 获取指定Panel名称和通道名称的通道配置
        /// </summary>
        public ChannelConfiguration? GetChannelConfiguration(string panelName, string channelName)
        {
            var config = GetPanelRuleConfiguration(panelName);
            return config?.Channels.FirstOrDefault(c => c.Name.Equals(channelName, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// 获取所有Panel规则配置 (现在只有一个)
        /// </summary>
        public List<PanelRuleConfiguration> GetAllPanelRuleConfigurations()
        {
            var list = new List<PanelRuleConfiguration>();
            if (_panelRuleConfiguration != null)
            {
                list.Add(_panelRuleConfiguration);
            }
            return list;
        }

        /// <summary>
        /// 保存Panel规则配置
        /// </summary>
        public bool SaveChanges()
        {
            if (_panelRuleConfiguration == null) return false;
            try
            {
                string jsonContent = JsonSerializer.Serialize(_panelRuleConfiguration, _jsonOptions);
                File.WriteAllText(_rulesFilePath, jsonContent);
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"错误: 保存规则文件时出错: {ex.Message}");
                return false;
            }
        }

        private PanelRuleConfiguration CreateDefaultConfig()
        {
             return new PanelRuleConfiguration
            {
                Name = "默认分析规则",
                Description = "系统默认分析规则配置",
                Version = "1.0",
                RuleGroups = new List<RuleGroup>
                {
                    new RuleGroup
                    {
                        Name = "默认规则组",
                        Priority = 1,
                        Rules = new List<Rule>
                        {
                            new Rule
                            {
                                Name = "默认阳性规则",
                                Condition = "Channel == 'FAM' && CtValue >= 10 && CtValue <= 35",
                                Action = "Result = Positive"
                            },
                            new Rule
                            {
                                Name = "默认阴性规则",
                                Condition = "Channel == 'FAM' && (CtValue < 10 || CtValue > 35 || CtValue == null)",
                                Action = "Result = Negative"
                            }
                        }
                    }
                },
                Channels = new List<ChannelConfiguration>
                {
                    new ChannelConfiguration
                    {
                        Name = "FAM",
                        Target = "Target1",
                        MinPositiveCt = 10,
                        MaxPositiveCt = 35
                    },
                    new ChannelConfiguration
                    {
                        Name = "VIC",
                        Target = "Target2",
                        MinPositiveCt = 10,
                        MaxPositiveCt = 35
                    }
                }
            };
        }
    }
} 