using System;
using System.Security.Cryptography;
using System.Text;

namespace docment_tools_client.Helpers
{
    /// <summary>
    /// 登录专用加密工具类（AES/ECB/PKCS5Padding）
    /// // 登录业务场景：密码加密（对应第2条）
    /// </summary>
    public static class LoginEncryptHelper
    {
        private const string SaltKey = "documentTools2026";

        /// <summary>
        /// AES加密用户密码
        /// </summary>
        /// <param name="plainText">明文密码</param>
        /// <returns>Base64编码的密文</returns>
        public static string EncryptPassword(string plainText)
        {
            if (string.IsNullOrEmpty(plainText)) return string.Empty;

            try
            {
                // 1. 生成密钥 (SHA-256取前16字节)
                byte[] key;
                using (var sha256 = SHA256.Create())
                {
                    byte[] hash = sha256.ComputeHash(Encoding.UTF8.GetBytes(SaltKey));
                    key = new byte[16];
                    Array.Copy(hash, 0, key, 0, 16);
                }

                // 2. AES加密配置 (ECB模式, PKCS7即PKCS5Padding in .NET)
                using (var aes = Aes.Create())
                {
                    aes.Key = key;
                    aes.Mode = CipherMode.ECB;
                    aes.Padding = PaddingMode.PKCS7;

                    using (var encryptor = aes.CreateEncryptor())
                    {
                        byte[] inputBytes = Encoding.UTF8.GetBytes(plainText);
                        byte[] encryptedBytes = encryptor.TransformFinalBlock(inputBytes, 0, inputBytes.Length);
                        return Convert.ToBase64String(encryptedBytes);
                    }
                }
            }
            catch (Exception ex)
            {
                LogHelper.Error($"密码加密失败：{ex.Message}");
                throw new Exception("密码处理异常，请重试");
            }
        }
    }
}
