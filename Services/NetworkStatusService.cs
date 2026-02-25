using System;
using System.Net.NetworkInformation;

namespace docment_tools_client.Services
{
    /// <summary>
    /// 网络状态监听服务
    /// // 登录业务场景：断网/联网状态同步逻辑（对应第7条）
    /// </summary>
    public static class NetworkStatusService
    {
        public static event Action<bool> NetworkAvailabilityChanged;

        public static bool IsConnected => NetworkInterface.GetIsNetworkAvailable();

        static NetworkStatusService()
        {
            NetworkChange.NetworkAvailabilityChanged += (s, e) =>
            {
                NetworkAvailabilityChanged?.Invoke(e.IsAvailable);
            };
        }
    }
}
