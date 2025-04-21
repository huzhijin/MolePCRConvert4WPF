using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using MolePCRConvert4WPF.Core.Services;
using Microsoft.Extensions.Logging;

namespace MolePCRConvert4WPF.Infrastructure.Services
{
    /// <summary>
    /// 用户设置服务实现，使用JSON文件持久化存储用户设置
    /// </summary>
    public class UserSettingsService : IUserSettingsService
    {
        private readonly string _settingsFilePath;
        private readonly ILogger<UserSettingsService>? _logger;
        private Dictionary<string, string> _settings;
        private readonly JsonSerializerOptions _jsonOptions = new JsonSerializerOptions { WriteIndented = true };

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="logger">日志记录器</param>
        public UserSettingsService(ILogger<UserSettingsService>? logger = null)
        {
            _logger = logger;
            _settingsFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "UserSettings.json");
            _settings = new Dictionary<string, string>();
            LoadSettingsFromFile();
        }

        /// <summary>
        /// 从文件加载所有设置
        /// </summary>
        private void LoadSettingsFromFile()
        {
            try
            {
                if (File.Exists(_settingsFilePath))
                {
                    string json = File.ReadAllText(_settingsFilePath);
                    var loadedSettings = JsonSerializer.Deserialize<Dictionary<string, string>>(json);
                    if (loadedSettings != null)
                    {
                        _settings = loadedSettings;
                        _logger?.LogInformation("成功从文件加载用户设置：{SettingsCount}项", _settings.Count);
                    }
                    else
                    {
                        _logger?.LogWarning("用户设置文件存在但反序列化结果为null");
                        _settings = new Dictionary<string, string>();
                    }
                }
                else
                {
                    _logger?.LogInformation("用户设置文件不存在，将使用空设置");
                    _settings = new Dictionary<string, string>();
                    SaveSettingsToFile(); // 创建空的设置文件
                }
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "加载用户设置时出错");
                _settings = new Dictionary<string, string>();
            }
        }

        /// <summary>
        /// 将设置保存到文件
        /// </summary>
        private bool SaveSettingsToFile()
        {
            try
            {
                string json = JsonSerializer.Serialize(_settings, _jsonOptions);
                File.WriteAllText(_settingsFilePath, json);
                _logger?.LogInformation("用户设置已保存到文件");
                return true;
            }
            catch (Exception ex)
            {
                _logger?.LogError(ex, "保存用户设置时出错");
                return false;
            }
        }

        /// <summary>
        /// 保存设置值
        /// </summary>
        public bool SaveSetting(string key, string value)
        {
            if (string.IsNullOrEmpty(key))
            {
                _logger?.LogWarning("尝试保存空键的设置");
                return false;
            }

            _settings[key] = value;
            _logger?.LogDebug("保存设置: {Key}={Value}", key, value);
            return SaveSettingsToFile();
        }

        /// <summary>
        /// 加载设置值
        /// </summary>
        public string LoadSetting(string key, string defaultValue = "")
        {
            if (string.IsNullOrEmpty(key))
            {
                _logger?.LogWarning("尝试加载空键的设置");
                return defaultValue;
            }

            if (_settings.TryGetValue(key, out string? value))
            {
                _logger?.LogDebug("加载设置: {Key}={Value}", key, value);
                return value;
            }

            _logger?.LogDebug("设置 {Key} 不存在，使用默认值 {DefaultValue}", key, defaultValue);
            return defaultValue;
        }

        /// <summary>
        /// 获取所有设置
        /// </summary>
        public Dictionary<string, string> GetAllSettings()
        {
            return new Dictionary<string, string>(_settings);
        }
    }
} 