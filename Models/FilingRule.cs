using System;

namespace docment_tools_client.Models
{
    /// <summary>
    /// 归档规则实体
    /// </summary>
    public class FilingRule
    {
        /// <summary>
        /// 规则名称
        /// </summary>
        public string RuleName { get; set; }

        /// <summary>
        /// 提取归档Key的方法（输入文件名，输出归档文件夹名）
        /// </summary>
        public Func<string, string> GetArchiveKey { get; set; }
    }
}