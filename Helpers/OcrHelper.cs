using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using docment_tools_client.Models;
using SixLabors.ImageSharp;
using Tesseract;

namespace docment_tools_client.Helpers
{
    /// <summary>
    /// OCR帮助类，专门处理OCR识别任务
    /// </summary>
    public class OcrHelper
    {
        /// <summary>
        /// 解析身份证OCR文本为键值对（精准匹配各字段）
        /// </summary>
        /// <param name="ocrText">OCR识别文本</param>
        /// <returns>身份证字段键值对</returns>
        public static Dictionary<string, string> ParseOcrTextToKeyValue(string ocrText)
        {
            var keyValueDict = new Dictionary<string, string>();
            if (string.IsNullOrWhiteSpace(ocrText))
            {
                keyValueDict.Add("识别状态", "无有效文本");
                return keyValueDict;
            }

            try
            {
                // 1. 匹配姓名（姓名：后接1-4个汉字）
                var nameMatch = Regex.Match(ocrText, @"姓名[:：]\s*([\u4e00-\u9fa5]{1,4})", RegexOptions.IgnoreCase);
                if (nameMatch.Success) keyValueDict.Add("姓名", nameMatch.Groups[1].Value.Trim());

                // 2. 匹配性别（性别：后接男/女）
                var genderMatch = Regex.Match(ocrText, @"性别[:：]\s*([男女])", RegexOptions.IgnoreCase);
                if (genderMatch.Success) keyValueDict.Add("性别", genderMatch.Groups[1].Value.Trim());

                // 3. 匹配民族（民族：后接1-4个汉字）
                var nationMatch = Regex.Match(ocrText, @"民族[:：]\s*([\u4e00-\u9fa5]{1,4})", RegexOptions.IgnoreCase);
                if (nationMatch.Success) keyValueDict.Add("民族", nationMatch.Groups[1].Value.Trim());

                // 4. 匹配出生日期（出生日期：后接YYYYMMDD/YYYY-MM-DD）
                var birthMatch = Regex.Match(ocrText, @"出生日期[:：]\s*(\d{4}[-/]?\d{2}[-/]?\d{2})", RegexOptions.IgnoreCase);
                if (birthMatch.Success)
                {
                    var birth = birthMatch.Groups[1].Value.Trim().Replace("/", "-");
                    keyValueDict.Add("出生日期", birth);
                }

                // 5. 匹配住址（住址：后接任意字符，直到身份证号前结束）
                var addressMatch = Regex.Match(ocrText, @"住址[:：]\s*(.+?)\s*公民身份号码", RegexOptions.Singleline | RegexOptions.IgnoreCase);
                if (addressMatch.Success) keyValueDict.Add("住址", addressMatch.Groups[1].Value.Trim());

                // 6. 匹配身份证号（18位，最后一位支持X/x）
                var idCardMatch = Regex.Match(ocrText, @"公民身份号码[:：]\s*([1-9]\d{5}(19|20)\d{2}((0[1-9])|(1[0-2]))(([0-2][1-9])|10|20|30|31)\d{3}[\dXx])", RegexOptions.IgnoreCase);
                if (idCardMatch.Success) keyValueDict.Add("公民身份号码", idCardMatch.Groups[1].Value.Trim().ToUpper());

                // 7. 补充识别状态
                keyValueDict.Add("识别状态", "成功");
                keyValueDict.Add("识别时间", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

                // 兜底：若未匹配到核心字段，保留原始文本
                if (keyValueDict.Count == 2) // 仅识别状态+时间
                {
                    keyValueDict.Add("原始文本", ocrText.Length > 500 ? ocrText.Substring(0, 500) : ocrText);
                }
            }
            catch (Exception ex)
            {
                keyValueDict.Add("识别状态", $"解析失败：{ex.Message}");
                keyValueDict.Add("原始文本", ocrText);
            }

            return keyValueDict;
        }

        public static Dictionary<string, string> ParseIdCardTextToKeyValue(string ocrText)
        {
            var keyValueDict = new Dictionary<string, string>();
            if (string.IsNullOrEmpty(ocrText)) return keyValueDict;

            var idCard = new IdCardModel();
            var lines = ocrText.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);

            foreach (var line in lines)
            {
                var trimmedLine = line.Trim();

                if (trimmedLine.StartsWith("姓名") || trimmedLine.Contains("姓名"))
                {
                    var match = Regex.Match(trimmedLine, @"姓名[：:\s]*([^\s\n\r]+)");
                    if (match.Success)
                    {
                        idCard.姓名 = match.Groups[1].Value.Trim();
                    }
                }
                else if (Regex.IsMatch(trimmedLine, @"^[^：:\s]*[：:]\s*[^\s\n\r]+$"))
                {
                    var parts = Regex.Split(trimmedLine, @"[：:]").Take(2).ToArray();
                    if (parts.Length == 2)
                    {
                        var field = parts[0].Trim();
                        var value = parts[1].Trim();

                        switch (field)
                        {
                            case "姓名":
                                idCard.姓名 = value;
                                break;
                            case "性别":
                                idCard.性别 = value;
                                break;
                            case "民族":
                                idCard.民族 = value;
                                break;
                            case "出生":
                                idCard.出生 = value;
                                break;
                            case "住址":
                                idCard.住址 = value;
                                break;
                            case "公民身份号码":
                                idCard.公民身份号码 = value;
                                break;
                            case "签发机关":
                                idCard.签发机关 = value;
                                break;
                            case "有效期限":
                                idCard.有效期限 = value;
                                break;
                        }
                    }
                }
                else
                {
                    idCard.姓名 = ExtractFieldFromText("姓名", idCard.姓名, trimmedLine);
                    idCard.性别 = ExtractFieldFromText("性别", idCard.性别, trimmedLine);
                    idCard.民族 = ExtractFieldFromText("民族", idCard.民族, trimmedLine);
                    idCard.出生 = ExtractFieldFromText("出生", idCard.出生, trimmedLine);
                    idCard.住址 = ExtractFieldFromText("住址", idCard.住址, trimmedLine);
                    idCard.公民身份号码 = ExtractFieldFromText("公民身份号码", idCard.公民身份号码, trimmedLine);
                    idCard.签发机关 = ExtractFieldFromText("签发机关", idCard.签发机关, trimmedLine);
                    idCard.有效期限 = ExtractFieldFromText("有效期限", idCard.有效期限, trimmedLine);
                }
            }

            return idCard.ToKeyValue();
        }

        /// <summary>
        /// 从文本中提取特定字段的值
        /// </summary>
        private static string ExtractFieldFromText(string fieldName, string fieldValue, string text)
        {
            if (!string.IsNullOrEmpty(fieldValue)) return fieldValue;

            var pattern = $@"{fieldName}[：:\s]*([^\s\n\r]+)";
            var match = Regex.Match(text, pattern);
            if (match.Success)
            {
                return match.Groups[1].Value.Trim();
            }
            return fieldValue;
        }
    }
}
