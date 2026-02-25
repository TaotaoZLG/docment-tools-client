using System;
using System.Collections.Generic;
using System.Text;
using Newtonsoft.Json;

namespace docment_tools_client.Models
{
    /// <summary>
    /// 用户信息模型
    /// </summary>
    public class UserInfo
    {
        /// <summary>
        /// 用户账号
        /// </summary>
        public string Account { get; set; } = string.Empty;

        /// <summary>
        /// 用户名
        /// </summary>
        public string UserName { get; set; } = string.Empty;

        /// <summary>
        /// 加密后的剩余额度
        /// </summary>
        public string EncryptedQuota { get; set; } = string.Empty;

        // 登录业务场景：Token存储（对应第3条）
        /// <summary>
        /// 登录Token
        /// </summary>
        public string Token { get; set; } = string.Empty;

        // 登录业务场景：单设备登录限制（对应第5条）
        /// <summary>
        /// 登录记录ID（用于状态同步与退出）
        /// </summary>
        public long LoginRecordId { get; set; }

        // 登录业务场景：用户状态持久化（对应第1条）
        /// <summary>
        /// 当前登录状态
        /// </summary>
        public UserStatus Status { get; set; } = UserStatus.EXIT;

        /// <summary>
        /// 剩余额度（运行时解密使用，不直接存储到JSON中，这里仅作为属性）
        /// </summary>
        [JsonIgnore]
        public decimal Quota { get; set; }

        // ==========================================
        // 登录业务新增字段 (对应需求)
        // ==========================================

        /// <summary>
        /// 登录用户ID
        /// </summary>
        public long UserId { get; set; }

        /// <summary>
        /// 登录用户昵称（用于登录后页面显示的用户名称信息，取代原UserName的展示作用）
        /// </summary>
        public string NickName 
        { 
            get => UserName; 
            set => UserName = value; 
        }

        // Quota 对应 UserBalance (用户额度)
        // 保持 Quota 命名兼容现有代码，但逻辑上它就是 UserBalance
        
        /// <summary>
        /// 用户单价（无需解密时使用，运行时属性）
        /// </summary>
        [JsonIgnore]
        public decimal UserPrice { get; set; }
        
        /// <summary>
        /// 加密后的用户单价 (存储与传输使用)
        /// </summary>
        public string EncryptedUserPrice { get; set; } = string.Empty;

        /// <summary>
        /// 上次登录时间
        /// </summary>
        public DateTime LastLoginTime { get; set; }

        /// <summary>
        /// 用户消息列表 (用于界面显示欢迎语等)
        /// </summary>
        public List<string> Messages { get; set; } = new List<string>();
    }
}
