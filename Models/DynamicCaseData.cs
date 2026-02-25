using System;
using System.Collections.Generic;
using System.Text;

namespace docment_tools_client.Models
{
    /// <summary>
    /// 动态案件数据模型（支持用户自定义Excel表头/Word占位符）
    /// </summary>
    public class DynamicCaseData
    {
        /// <summary>
        /// 案件唯一标识
        /// </summary>
        public string CaseId { get; set; } = Guid.NewGuid().ToString();

        /// <summary>
        /// 表头与数据的键值对（兼容任意Excel表头）
        /// </summary>
        public Dictionary<string, object> CaseInfo { get; set; } = new Dictionary<string, object>();

        /// <summary>
        /// 快速获取指定表头对应的数据（避免键不存在报错）
        /// </summary>
        /// <param name="header">Excel表头/Word占位符</param>
        /// <returns>对应的数据，无对应键返回空字符串</returns>
        public string GetValue(string header)
        {
            if (CaseInfo.TryGetValue(header, out object? value))
            {
                return value?.ToString() ?? string.Empty;
            }
            return string.Empty;
        }

        /// <summary>
        /// 是否已消费（扣费标记）
        /// </summary>
        public bool IsConsumed { get; set; } = false;
    }
}
