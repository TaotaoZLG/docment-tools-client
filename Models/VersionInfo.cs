using System;
using System.Collections.Generic;
using System.Text;

namespace docment_tools_client.Models
{
    /// <summary>
    /// 版本信息模型
    /// </summary>
    public class VersionInfo
    {
        /// <summary>
        /// 版本号（如1.0.0.0）
        /// </summary>
        public string Version { get; set; } = "1.0.0.0";

        /// <summary>
        /// 更新内容
        /// </summary>
        public string UpdateContent { get; set; } = string.Empty;

        /// <summary>
        /// 下载地址
        /// </summary>
        public string DownloadUrl { get; set; } = string.Empty;

        /// <summary>
        /// 是否强制更新
        /// </summary>
        public bool IsForceUpdate { get; set; } = false;

        /// <summary>
        /// 发布时间
        /// </summary>
        public DateTime PublishTime { get; set; } = DateTime.Now;
    }
}
