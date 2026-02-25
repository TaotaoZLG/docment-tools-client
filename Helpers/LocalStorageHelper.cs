using docment_tools_client.Models;
using docment_tools_client.Helpers;
using Newtonsoft.Json;
using System;
using System.IO;

namespace docment_tools_client.Helpers
{
    /// <summary>
    /// 本地配置存储工具类（基于JSON文件，存储用户信息、应用配置）
    /// </summary>
    public static class LocalStorageHelper
    {
        /// <summary>
        /// 应用本地存储目录（%AppData%\Local\DocumentToolsClient）
        /// </summary>
        private static readonly string _appDataDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "DocumentToolsClient");

        /// <summary>
        /// 用户信息配置文件路径
        /// </summary>
        private static readonly string _userConfigPath;

        /// <summary>
        /// 应用自带（%AppData%\Local\DocumentToolsClient\SystemTemplates）
        /// </summary>
        public static readonly string _systemTemplateDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "DocumentToolsClient",
            "SystemTemplates");

        /// <summary>
        /// 静态构造函数（初始化目录和文件）
        /// </summary>
        static LocalStorageHelper()
        {
            // 创建目录（不存在则创建）
            if (!Directory.Exists(_appDataDir))
            {
                Directory.CreateDirectory(_appDataDir);
            }

            // 初始化用户配置文件路径
            _userConfigPath = Path.Combine(_appDataDir, "UserConfig.json");

            // 初始化系统模板文件路径
            if (!Directory.Exists(_systemTemplateDir))
            {
                Directory.CreateDirectory(_systemTemplateDir);
                LogHelper.Info($"系统模板目录已创建：{_systemTemplateDir}");
            }
        }

        /// <summary>
        /// 保存用户信息到本地JSON文件（整体加密存储）
        /// </summary>
        public static void SaveUserInfo(UserInfo userInfo)
        {
            if (userInfo == null) return;

            try
            {
                // 将UserInfo对象整体加密为密文字符串
                var encryptedUserStr = EncryptHelper.EncryptObject(userInfo);
                // 将密文字符串写入本地文件
                File.WriteAllText(_userConfigPath, encryptedUserStr, System.Text.Encoding.UTF8);
            }
            catch (Exception ex)
            {
                LogHelper.Error($"保存用户信息失败：{ex.Message}");
            }
        }

        /// <summary>
        /// 从本地JSON文件读取用户信息（先解密再反序列化）
        /// </summary>
        public static UserInfo? GetUserInfo()
        {
            if (!File.Exists(_userConfigPath))
            {
                return null;
            }

            try
            {
                // 1. 读取本地文件中的密文字符串
                var encryptedUserStr = File.ReadAllText(_userConfigPath, System.Text.Encoding.UTF8);
                // 2. 解密并反序列化为UserInfo对象
                var user = EncryptHelper.DecryptObject<UserInfo>(encryptedUserStr);
                
                // 3. 解密敏感字段 (额度和单价)
                if (user != null)
                {
                    user.Quota = EncryptHelper.DecryptQuota(user.EncryptedQuota);
                    user.UserPrice = EncryptHelper.DecryptQuota(user.EncryptedUserPrice);
                }
                
                return user;
            }
            catch (Exception ex)
            {
                LogHelper.Error($"读取用户信息失败：{ex.Message}");
                // 如果读取失败（可能是密钥变更或文件损坏），建议删除文件需重新登录
                DeleteUserConfig();
                return null;
            }
        }
        
        /// <summary>
        /// 删除用户本地配置文件
        /// </summary>
        public static void DeleteUserConfig()
        {
             if (File.Exists(_userConfigPath))
             {
                 try
                 {
                     File.Delete(_userConfigPath);
                     LogHelper.Info("本地用户配置文件已删除");
                 }
                 catch {}
             }
        }

        /// <summary>
        /// 保存通用配置（键值对）
        /// </summary>
        public static void SaveConfig<T>(string key, T value)
        {
            var configPath = Path.Combine(_appDataDir, "AppConfig.json");
            var configDict = new Dictionary<string, object>();

            // 读取现有配置
            if (File.Exists(configPath))
            {
                var json = File.ReadAllText(configPath);
                configDict = JsonConvert.DeserializeObject<Dictionary<string, object>>(json) ?? new Dictionary<string, object>();
            }

            // 更新配置
            if (configDict.ContainsKey(key))
            {
                configDict[key] = value;
            }
            else
            {
                configDict.Add(key, value);
            }

            // 写入文件
            var newJson = JsonConvert.SerializeObject(configDict, Formatting.Indented);
            File.WriteAllText(configPath, newJson, System.Text.Encoding.UTF8);
        }

        /// <summary>
        /// 获取通用配置
        /// </summary>
        public static T? GetConfig<T>(string key)
        {
            var configPath = Path.Combine(_appDataDir, "AppConfig.json");
            if (!File.Exists(configPath))
            {
                return default;
            }

            var json = File.ReadAllText(configPath);
            var configDict = JsonConvert.DeserializeObject<Dictionary<string, object>>(json) ?? new Dictionary<string, object>();

            if (configDict.ContainsKey(key))
            {
                return JsonConvert.DeserializeObject<T>(configDict[key].ToString() ?? string.Empty);
            }

            return default;
        }
    }
}
