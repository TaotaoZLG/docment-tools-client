using Newtonsoft.Json;
using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace docment_tools_client.Helpers
{
    /// <summary>
    /// AES加密解密工具类（适配.NET 10，用于加密用户信息和本地额度）
    /// </summary>
    public static class EncryptHelper
    {
        /// <summary>
        /// 加密密钥（生产环境从安全渠道获取，此处为演示固定值，长度16/24/32）
        /// </summary>
        private static readonly string _key = "DocmentTools_SecretKey_123456789"; // 32位

        /// <summary>
        /// 加密向量（长度必须16位）
        /// </summary>
        private static readonly string _iv = "DocTools_IV_5678"; // 16位


        /// <summary>
        /// 验证密钥/向量长度
        /// </summary>
        static EncryptHelper()
        {
            if (_key.Length != 32)
            {
                throw new ArgumentException("AES密钥必须为32位字符串");
            }
            if (_iv.Length != 16)
            {
                throw new ArgumentException("AES初始化向量必须为16位字符串");
            }
        }

        /// <summary>
        /// AES加密字符串
        /// </summary>
        /// <param name="plainText">明文</param>
        /// <returns>加密后的Base64字符串</returns>
        public static string AesEncrypt(string plainText)
        {
            if (string.IsNullOrEmpty(plainText)) return string.Empty;

            using var aes = Aes.Create();
            aes.Key = Encoding.UTF8.GetBytes(_key);
            aes.IV = Encoding.UTF8.GetBytes(_iv);
            aes.Mode = CipherMode.CBC;
            aes.Padding = PaddingMode.PKCS7;

            // 创建加密器
            using var encryptor = aes.CreateEncryptor(aes.Key, aes.IV);
            using var ms = new MemoryStream();
            using var cs = new CryptoStream(ms, encryptor, CryptoStreamMode.Write);

            // 写入明文并加密
            var plainBytes = Encoding.UTF8.GetBytes(plainText);
            cs.Write(plainBytes, 0, plainBytes.Length);
            cs.FlushFinalBlock();

            // 转换为Base64字符串返回
            return Convert.ToBase64String(ms.ToArray());
        }

        /// <summary>
        /// AES解密字符串
        /// </summary>
        /// <param name="cipherText">加密后的Base64字符串</param>
        /// <returns>明文</returns>
        public static string AesDecrypt(string cipherText)
        {
            if (string.IsNullOrEmpty(cipherText)) return string.Empty;

            using var aes = Aes.Create();
            aes.Key = Encoding.UTF8.GetBytes(_key);
            aes.IV = Encoding.UTF8.GetBytes(_iv);
            aes.Mode = CipherMode.CBC;
            aes.Padding = PaddingMode.PKCS7;

            // 创建解密器
            using var decryptor = aes.CreateDecryptor(aes.Key, aes.IV);
            using var ms = new MemoryStream();
            using var cs = new CryptoStream(ms, decryptor, CryptoStreamMode.Write);

            // 转换加密字符串为字节数组
            var cipherBytes = Convert.FromBase64String(cipherText);
            cs.Write(cipherBytes, 0, cipherBytes.Length);
            cs.FlushFinalBlock();

            // 转换为明文返回
            return Encoding.UTF8.GetString(ms.ToArray());
        }

        /// <summary>
        /// 额度加密（专门用于加密decimal类型的额度）
        /// </summary>
        public static string EncryptQuota(decimal quota)
        {
            return AesEncrypt(quota.ToString("F2"));
        }

        /// <summary>
        /// 额度解密（专门用于解密获取decimal类型的额度）
        /// </summary>
        public static decimal DecryptQuota(string encryptedQuota)
        {
            var plainText = AesDecrypt(encryptedQuota);
            if (decimal.TryParse(plainText, out var quota))
            {
                return quota;
            }
            return 0.00m;
        }

        /// <summary>
        /// 复杂对象序列化后AES加密（用于整体加密UserInfo等对象）
        /// </summary>
        /// <typeparam name="T">对象类型</typeparam>
        /// <param name="obj">待加密对象</param>
        /// <returns>加密后的Base64字符串</returns>
        public static string EncryptObject<T>(T obj)
        {
            if (obj == null) return string.Empty;

            // 1. 先将对象序列化为JSON字符串（明文）
            var jsonStr = JsonConvert.SerializeObject(obj, Formatting.Indented);
            // 2. 再对JSON字符串进行AES加密
            return AesEncrypt(jsonStr);
        }

        /// <summary>
        /// AES解密后反序列化为复杂对象（用于读取加密的UserInfo等对象）
        /// </summary>
        /// <typeparam name="T">对象类型</typeparam>
        /// <param name="cipherText">加密后的Base64字符串</param>
        /// <returns>反序列化后的对象</returns>
        public static T? DecryptObject<T>(string cipherText)
        {
            if (string.IsNullOrEmpty(cipherText)) return default;

            // 1. 先对密文进行AES解密，得到JSON字符串
            var jsonStr = AesDecrypt(cipherText);
            // 2. 再将JSON字符串反序列化为对象
            return JsonConvert.DeserializeObject<T>(jsonStr);
        }

    }
}
