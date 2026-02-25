namespace docment_tools_client.Models
{
    // 登录业务场景：用户状态管理（对应第1条）
    public enum UserStatus
    {
        /// <summary>
        /// 登录在线
        /// </summary>
        ONLINE,

        /// <summary>
        /// 登录离线（断网保持登录）
        /// </summary>
        OFFLINE,

        /// <summary>
        /// 已退出
        /// </summary>
        EXIT
    }

    // 登录业务场景：设备唯一标识辅助（对应第5条）
    public static class DeviceInfoModel
    {
        public static string CurrentDeviceIdentifier { get; set; } = string.Empty;
    }
}
