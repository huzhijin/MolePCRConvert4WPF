using System;
using System.Collections.Generic;

namespace MolePCRConvert4WPF.Core.Services
{
    /// <summary>
    /// 用户设置服务接口，负责持久化存储和加载用户设置
    /// </summary>
    public interface IUserSettingsService
    {
        /// <summary>
        /// 保存设置值
        /// </summary>
        /// <param name="key">设置键</param>
        /// <param name="value">设置值</param>
        /// <returns>是否保存成功</returns>
        bool SaveSetting(string key, string value);

        /// <summary>
        /// 加载设置值
        /// </summary>
        /// <param name="key">设置键</param>
        /// <param name="defaultValue">默认值（如果设置不存在）</param>
        /// <returns>设置值</returns>
        string LoadSetting(string key, string defaultValue = "");

        /// <summary>
        /// 获取所有设置
        /// </summary>
        /// <returns>所有设置的键值对</returns>
        Dictionary<string, string> GetAllSettings();
    }
} 