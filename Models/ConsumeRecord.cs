using System;
using System.Collections.Generic;
using System.Text;

namespace docment_tools_client.Models
{
    public class ConsumeRecord
    {
        /// <summary>
        /// 记录唯一标识（防止重复同步，使用GUID）
        /// </summary>
        public string RecordId { get; set; } = Guid.NewGuid().ToString();

        /// <summary>
        /// 用户账号
        /// </summary>
        public string UserAccount { get; set; } = string.Empty;

        /// <summary>
        /// 消费案件条数
        /// </summary>
        public int CaseCount { get; set; }

        /// <summary>
        /// 消费金额（案件条数 * 单价）
        /// </summary>
        public decimal Amount { get; set; }

        /// <summary>
        /// 消费时间
        /// </summary>
        public DateTime ConsumeTime { get; set; } = DateTime.Now;

        /// <summary>
        /// 是否已同步到后台
        /// </summary>
        public bool IsSynced { get; set; } = false;

        /// <summary>
        /// 文书类型（起诉状/保全申请书等）
        /// </summary>
        public string DocumentType { get; set; } = string.Empty;

        /// <summary>
        /// 文档的案件填充数量
        /// </summary>
        public int DocumentCaseCount { get; set; }

        /// <summary>
        /// 输出文件路径
        /// </summary>
        public string FilePath { get; set; } = string.Empty;

    }
}
