using System;
using System.Collections.Generic;
using System.Text;

namespace docment_tools_client.Models
{
    /// <summary>
    /// 案件数据模型（对应Excel数据和Word模板填充字段）
    /// </summary>
    public class CaseData
    {
        /// <summary>
        /// 姓名
        /// </summary>
        public string Name { get; set; } = string.Empty;

        /// <summary>
        /// 身份证号
        /// </summary>
        public string IdCard { get; set; } = string.Empty;

        /// <summary>
        /// 案由
        /// </summary>
        public string CaseReason { get; set; } = string.Empty;

        /// <summary>
        /// 案件金额（元）
        /// </summary>
        public decimal Amount { get; set; }

        /// <summary>
        /// 案件日期
        /// </summary>
        public DateTime CaseDate { get; set; } = DateTime.Now;

        /// <summary>
        /// 是否需要附件表格
        /// </summary>
        public bool NeedAttachmentTable { get; set; }

        /// <summary>
        /// 附件表格数据（一拖多场景）
        /// </summary>
        public List<CaseData> AttachmentData { get; set; } = new List<CaseData>();
    }
}
