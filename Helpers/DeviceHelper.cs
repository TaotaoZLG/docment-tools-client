using System;
using System.IO;
using System.IO.IsolatedStorage;
using System.Management;
using System.Security.Cryptography;
using System.Text;

namespace docment_tools_client.Helpers
{
    /// <summary>
    /// 设备唯一标识帮助类（优化版）
    /// 策略：优先读取本地缓存 -> 其次通过主板+BIOS生成 -> 兜底使用GUID
    /// </summary>
    public static class DeviceHelper
    {
        private static string _cachedDeviceId = string.Empty;
        private const string CacheFileName = "device_identity.dat";

        /// <summary>
        /// 获取设备唯一标识
        /// </summary>
        public static string GetDeviceIdentifier()
        {
            // 1. 内存缓存
            if (!string.IsNullOrEmpty(_cachedDeviceId)) return _cachedDeviceId;

            try
            {
                // 2. 持久化存储缓存
                _cachedDeviceId = GetPersistedDeviceId();
                if (!string.IsNullOrEmpty(_cachedDeviceId)) return _cachedDeviceId;

                // 3. 生成新硬件特征码
                string hardwareCode = GetHardwareCharacteristics();
                
                // 4. Hash处理生成最终ID
                string finalId = ComputeHash(hardwareCode);
                
                _cachedDeviceId = finalId;

                // 5. 保存
                SaveDeviceId(_cachedDeviceId);
            }
            catch (Exception ex)
            {
                LogHelper.Error($"生成设备ID异常：{ex.Message}");
                // 6. 异常兜底
                if (string.IsNullOrEmpty(_cachedDeviceId))
                {
                    _cachedDeviceId = Guid.NewGuid().ToString("N");
                }
            }

            return _cachedDeviceId;
        }

        /// <summary>
        /// 获取硬件特征（主板序列号 + BIOS UUID）
        /// </summary>
        private static string GetHardwareCharacteristics()
        {
            var sb = new StringBuilder();
            
            // 优先顺序：BIOS UUID > 主板序列号 > 处理器ID
            string uuid = GetWmiInfo("Win32_ComputerSystemProduct", "UUID");
            string boardSerial = GetWmiInfo("Win32_BaseBoard", "SerialNumber");
            string cpuId = GetWmiInfo("Win32_Processor", "ProcessorId");

            // 有效性检查并拼接
            if (IsValidHardwareValue(uuid)) sb.Append(uuid);
            if (IsValidHardwareValue(boardSerial)) sb.Append(boardSerial);
            
            // 如果都获取失败，使用CPU ID和机器名补充
            if (sb.Length < 10) 
            {
                if (IsValidHardwareValue(cpuId)) sb.Append(cpuId);
                sb.Append(Environment.MachineName);
            }
            
            // 如果还是没有任何信息（极端情况），生成一个随机数防止空
            if (sb.Length == 0)
            {
                return Guid.NewGuid().ToString();
            }

            return sb.ToString();
        }

        private static string GetWmiInfo(string cls, string prop)
        {
            try
            {
                using (var searcher = new ManagementObjectSearcher($"SELECT {prop} FROM {cls}"))
                {
                    foreach (var item in searcher.Get())
                    {
                        string val = item[prop]?.ToString();
                        if (IsValidHardwareValue(val)) return val.Trim();
                    }
                }
            }
            catch (Exception ex)
            {
                LogHelper.Warn($"获取硬件信息({cls})失败: {ex.Message}");
            }
            return string.Empty;
        }

        private static bool IsValidHardwareValue(string val)
        {
            if (string.IsNullOrWhiteSpace(val)) return false;
            // 排除常见无效值
            string v = val.Trim().ToUpper();
            return v != "NONE" && v != "UNKNOWN" && v != "DEFAULT STRING" && v != "0" && v != "TO BE FILLED BY O.E.M.";
        }

        private static string ComputeHash(string input)
        {
            using (var sha256 = SHA256.Create())
            {
                byte[] bytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(input + "-Salt-DocTools"));
                StringBuilder builder = new StringBuilder();
                foreach (byte b in bytes)
                {
                    builder.Append(b.ToString("x2"));
                }
                return builder.ToString();
            }
        }

        private static void SaveDeviceId(string deviceId)
        {
            try
            {
                string encrypted = EncryptHelper.AesEncrypt(deviceId);
                
                // 使用 IsolatedStorage 替代本地文件路径，避免权限问题
                using (var isoStore = IsolatedStorageFile.GetUserStoreForAssembly())
                {
                    using (var stream = new IsolatedStorageFileStream(CacheFileName, FileMode.Create, isoStore))
                    using (var writer = new StreamWriter(stream))
                    {
                        writer.Write(encrypted);
                    }
                }
            }
            catch (Exception ex)
            {
                LogHelper.Error($"存储设备ID失败：{ex.Message}");
            }
        }

        private static string GetPersistedDeviceId()
        {
            try
            {
                using (var isoStore = IsolatedStorageFile.GetUserStoreForAssembly())
                {
                    if (isoStore.FileExists(CacheFileName))
                    {
                        using (var stream = new IsolatedStorageFileStream(CacheFileName, FileMode.Open, isoStore))
                        using (var reader = new StreamReader(stream))
                        {
                            string encrypted = reader.ReadToEnd();
                            if (!string.IsNullOrEmpty(encrypted))
                            {
                                return EncryptHelper.AesDecrypt(encrypted);
                            }
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                LogHelper.Warn($"读取本地设备ID失败：{ex.Message}");
            }
            return string.Empty;
        }

        // 兼容旧方法的存根（如果别的地方用到了这些私有方法，虽然DeviceHelper是internal/public，但private方法只需内部兼容）
        private static string GetOrSetFallbackDeviceId() 
        {
            // 复用主逻辑，因为主逻辑已包含生成和保存
             return GetDeviceIdentifier(); 
        }
    }
}
