using System;
using System.Collections.Generic;
using System.Net.NetworkInformation;
using System.Text;

namespace docment_tools_client.Helpers
{
    /// <summary>
    /// 网络状态检测工具类
    /// </summary>
    public static class NetworkHelper
    {
        /// <summary>
        /// 检测是否有网络连接（通用检测）
        /// </summary>
        public static bool IsNetworkAvailable()
        {
            try
            {
                // 检测是否有可用的网络连接（排除断开的网卡）
                return NetworkInterface.GetIsNetworkAvailable() &&
                       NetworkInterface.GetAllNetworkInterfaces()
                           .Any(ni => ni.OperationalStatus == OperationalStatus.Up &&
                                      ni.NetworkInterfaceType != NetworkInterfaceType.Loopback &&
                                      ni.NetworkInterfaceType != NetworkInterfaceType.Tunnel);
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// 检测是否能访问指定服务器（用于检测后台接口是否可达）
        /// </summary>
        /// <param name="host">服务器地址（如api.legaldoc.com）</param>
        public static bool IsServerReachable(string host = "www.baidu.com")
        {
            if (!IsNetworkAvailable()) return false;

            try
            {
                using var ping = new Ping();
                // 发送ping请求（超时3000毫秒）
                var reply = ping.Send(host, 3000);
                // 验证ping响应是否成功
                return reply != null && reply.Status == IPStatus.Success;
            }
            catch (Exception)
            {
                return false;
            }
        }
    }
}
