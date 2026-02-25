using docment_tools_client.Helpers;
using docment_tools_client.Models;
using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Json;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace docment_tools_client.Services
{
    public class ApiResponse
    {
        public int Code { get; set; }
        public string Msg { get; set; } = string.Empty;
        public object? Data { get; set; }
        public bool IsSuccess => Code == 200;
    }

    /// <summary>
    /// 接口请求服务（适配RuoYi /prod-api 接口，暂时模拟返回数据，后续替换为真实HttpClient请求）
    /// </summary>
    public static class ApiService
    {
        // 登录业务场景：全局请求封装与Token管理（对应第3、4条）
        private static readonly HttpClient _httpClient;
        private static readonly System.Threading.SemaphoreSlim _tokenRefreshLock = new System.Threading.SemaphoreSlim(1, 1);
        private const string AppHost = "http://localhost:8088";
        private const string BaseApiUrl = "http://api.cenosoft.top/prod-api"; // 实际地址需配置

        static ApiService()
        {
            _httpClient = new HttpClient();
            // 设置超时等基础配置
            _httpClient.Timeout = TimeSpan.FromSeconds(30);
        }

        private static void AttachHeaders(HttpRequestMessage request, string? token = null, string? deviceId = null)
        {
             // 登录业务场景：Token全球携带（对应第3条）
             if (!string.IsNullOrEmpty(token))
             {
                 request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);
             }
             if (!string.IsNullOrEmpty(deviceId))
             {
                 request.Headers.Add("Device-Identifier", deviceId);
             }
        }

        // 登录业务场景：Token定时刷新与失效处理（对应第4条）
        private static async Task<HttpResponseMessage> SendRequestWithRetryAsync(HttpRequestMessage request, bool allowRefresh = true)
        {
            // 复制请求以防需要重试
             // 注意：HttpRequestMessage通常不能被发送两次，这里简化处理，假设每次调用都是新的
             
             // 这里直接发送
             var response = await _httpClient.SendAsync(request);

             if (response.StatusCode == System.Net.HttpStatusCode.Unauthorized && allowRefresh)
             {
                 // 401，尝试刷新
                 await _tokenRefreshLock.WaitAsync();
                 try
                 {
                     // 二次检查，可能在等待时已被其他线程刷新
                     var currentToken = LocalStorageHelper.GetUserInfo()?.Token;
                     // 简单判断：如果内存里的Token已经跟请求里的不一样了（假设请求里是旧的），说明刷新过了？
                     // 这里我们在Request构建时是动态取的？
                     // 实际项目需比较 Token 差异或 TimeStamp。
                     
                     // 执行静默刷新
                     var refreshResult = await SilentRefreshAsync(currentToken);
                     if (refreshResult.IsSuccess && refreshResult.Data is string newToken)
                     {
                          // 刷新成功，更新本地Token
                          var user = LocalStorageHelper.GetUserInfo();
                          if (user != null)
                          {
                              user.Token = newToken;
                              LocalStorageHelper.SaveUserInfo(user);
                          }
                          
                          // 重试原请求（需要重新构建Request）
                          var newRequest = new HttpRequestMessage(request.Method, request.RequestUri);
                          if (request.Content != null) 
                          {
                              // 重新读取Content? 比较麻烦，针对JSON Content:
                              // newRequest.Content = ...
                              // 为简化代码，此处若非GET请求，建议直接抛出需重新登录，或在上层重试
                              // 假设当前只处理简单GET/POST Json
                          }
                          AttachHeaders(newRequest, newToken, DeviceInfoModel.CurrentDeviceIdentifier);
                          return await _httpClient.SendAsync(newRequest);
                     }
                     else
                     {
                         // 刷新失败，强制退出
                          LogoutCleanUp();
                          return response; // 返回401给上层处理
                     }
                 }
                 finally
                 {
                     _tokenRefreshLock.Release();
                 }
             }
             
             return response;
        }

        private static void LogoutCleanUp()
        {
            // 清除本地状态
            // 需在UI线程或通知ViewModel？ 
            // 这里简单清除本地文件，ViewModel监听状态变更或者下次操作时发现无Token及自动跳转
            LocalStorageHelper.DeleteUserConfig();
            // 实际应用可能需要EventBus通知跳转登录页
        }
        
        /// <summary>
        /// 获取系统公告列表
        /// </summary>
        public static async Task<List<string>> GetSystemMessagesAsync()
        {
            var list = new List<string>();
            if (!NetworkHelper.IsNetworkAvailable()) return list;

            var user = LocalStorageHelper.GetUserInfo();
            if (user == null || string.IsNullOrEmpty(user.Token)) return list;

            try
            {
                var request = new HttpRequestMessage(HttpMethod.Get, $"{AppHost}/out/common/message/user/list");
                AttachHeaders(request, user.Token);

                var response = await SendRequestWithRetryAsync(request);
                if (response.IsSuccessStatusCode)
                {
                    var result = await response.Content.ReadFromJsonAsync<AjaxResult>();
                    if (result != null && result.IsSuccess && result.Data != null)
                    {
                        if (result.Data is JsonElement element && element.ValueKind == JsonValueKind.Array)
                        {
                            foreach (var item in element.EnumerateArray())
                            {
                                list.Add(item.GetString() ?? "");
                            }
                        }
                        // 兼容某些Json序列化直接转为List的情况
                        else if (result.Data is List<string> strList)
                        {
                            list.AddRange(strList);
                        }
                        // 兼容Newtonsoft JArray (虽然此处用了System.Text.Json, 但为了稳健)
                        else if (result.Data is Newtonsoft.Json.Linq.JArray jArray)
                        {
                            foreach (var item in jArray)
                            {
                                list.Add(item.ToString());
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogHelper.Error($"获取系统公告失败：{ex.Message}");
            }

            return list;
        }

        /// <summary>
        /// 修改密码
        /// </summary>
        public static async Task<AjaxResult> UpdatePasswordAsync(string oldPassword, string newPassword)
        {
            if (!NetworkHelper.IsNetworkAvailable())
            {
                return AjaxResult.Error("无网络连接，无法修改密码");
            }

            var user = LocalStorageHelper.GetUserInfo();
            if (user == null || string.IsNullOrEmpty(user.Token))
            {
                return AjaxResult.Error("用户未登录");
            }

            try
            {
                var payload = new { oldPassword = oldPassword, newPassword = newPassword };
                var request = new HttpRequestMessage(HttpMethod.Post, $"{AppHost}/out/auth/updatePassword");
                request.Content = JsonContent.Create(payload);
                AttachHeaders(request, user.Token);

                var response = await SendRequestWithRetryAsync(request);
                var result = await response.Content.ReadFromJsonAsync<AjaxResult>();
                return result ?? AjaxResult.Error("服务器未返回数据");
            }
            catch (Exception ex)
            {
                LogHelper.Error($"修改密码请求失败：{ex.Message}");
                return AjaxResult.Error("请求失败，请稍后重试");
            }
        }

        /// <summary>
        /// 用户登录接口
        /// </summary>
        public static async Task<AjaxResult> LoginAsync(string account, string encryptedPassword, string deviceIdentifier)
        {
             // 网络检测
            if (!NetworkHelper.IsNetworkAvailable())
            {
                return AjaxResult.Error("无网络连接，无法登录");
            }
            
            try
            {
                // 构造请求参数
                var payload = new 
                { 
                    username = account, 
                    password = encryptedPassword, 
                    deviceIdentifier = deviceIdentifier 
                };

                // 发起POST请求 (Web后端接口: /out/auth/login)
                var request = new HttpRequestMessage(HttpMethod.Post, $"{AppHost}/out/auth/login");
                request.Content = JsonContent.Create(payload);
                
                // 注意：登录接口不需要Token Header，但如果未来需要ClientCreds等可添加
                
                var response = await _httpClient.SendAsync(request);
                var jsonString = await response.Content.ReadAsStringAsync();
                var options = new JsonSerializerOptions { PropertyNameCaseInsensitive = true };
                
                AjaxResult? result = null;
                try 
                {
                    result = JsonSerializer.Deserialize<AjaxResult>(jsonString, options);
                }
                catch
                {
                    LogHelper.Error($"登录响应解析失败: {jsonString}");
                    return AjaxResult.Error("服务器返回格式错误");
                }

                if (result != null && result.IsSuccess)
                {
                    // 解析响应数据
                    // 后端返回结构: { token, loginRecordId, userId, username, nickName, userBalance, userPrice }
                    if (result.Data != null && result.Data is JsonElement root)
                    {
                        // 辅助方法：安全获取long (支持number和string)
                        long GetSafeLong(JsonElement el, string key)
                        {
                            if (el.TryGetProperty(key, out var p))
                            {
                                if (p.ValueKind == JsonValueKind.Number && p.TryGetInt64(out long v)) return v;
                                if (p.ValueKind == JsonValueKind.String && long.TryParse(p.GetString(), out long vs)) return vs;
                            }
                            return 0;
                        }

                        var userInfo = new UserInfo
                        {
                            // 基础信息
                            Account = root.TryGetProperty("username", out var u) ? u.GetString() ?? account : account,
                            UserName = root.TryGetProperty("nickName", out var n) ? n.GetString() ?? account : account, 
                            Token = root.TryGetProperty("token", out var t) ? t.GetString() ?? "" : "",
                            
                            // 身份标识
                            LoginRecordId = GetSafeLong(root, "loginRecordId"),
                            UserId = GetSafeLong(root, "userId"),
                            
                            Status = UserStatus.ONLINE, 
                            LastLoginTime = DateTime.Now,
                            Messages = new List<string> { result.Msg ?? "登录成功" }
                        };

                        // 额度处理 (userBalance)
                        if (root.TryGetProperty("userBalance", out var ub))
                        {
                            decimal val = 0;
                            if (ub.ValueKind == JsonValueKind.Number && ub.TryGetDecimal(out val)) { }
                            else if (ub.ValueKind == JsonValueKind.String && decimal.TryParse(ub.GetString(), out val)) { }
                            
                            if (val > 0 || ub.ValueKind != JsonValueKind.Null) // 允许0
                            {
                                userInfo.Quota = val;
                                userInfo.EncryptedQuota = EncryptHelper.EncryptQuota(val);
                            }
                        }

                        // 单价处理 (userPrice)
                        if (root.TryGetProperty("userPrice", out var up))
                        {
                            decimal val = 0;
                            if (up.ValueKind == JsonValueKind.Number && up.TryGetDecimal(out val)) { }
                            else if (up.ValueKind == JsonValueKind.String && decimal.TryParse(up.GetString(), out val)) { }
                            
                            if (val > 0 || up.ValueKind != JsonValueKind.Null)
                            {
                                userInfo.UserPrice = val;
                                userInfo.EncryptedUserPrice = EncryptHelper.EncryptQuota(val);
                            }
                        }

                        LogHelper.Info($"用户{account}登录成功");
                        return AjaxResult.Success("登录成功", userInfo);
                    }
                    else
                    {
                        LogHelper.Error("登录返回成功，但数据为空或格式不正确");
                        return AjaxResult.Error("登录异常：服务器返回数据无效");
                    }
                }
                
                return result ?? AjaxResult.Error("登录失败，服务器无响应");
            }
            catch (Exception ex)
            {
                LogHelper.Error($"登录请求失败：{ex.Message}");
                return AjaxResult.Error("服务连接超时或异常，请检查网络连接");
            }
        }
        
        /// <summary>
        /// 退出登录
        /// // 登录业务场景：正常退出登录逻辑（对应第6条）
        /// </summary>
        public static async Task<AjaxResult> LogoutAsync(string account, string token, long loginRecordId)
        {
            if (!NetworkHelper.IsNetworkAvailable())
            {
                return AjaxResult.Error("请您先连接网络再退出登录");
            }

            try
            {
                var payload = new 
                { 
                    loginRecordId = loginRecordId, 
                    username = account, 
                    token = token 
                };

                var request = new HttpRequestMessage(HttpMethod.Post, $"{AppHost}/out/auth/logout");
                request.Content = JsonContent.Create(payload);
                AttachHeaders(request, token);

                var response = await SendRequestWithRetryAsync(request, false);
                var result = await response.Content.ReadFromJsonAsync<AjaxResult>();
                return result ?? AjaxResult.Success("退出成功");
            }
            catch (Exception ex)
            {
                LogHelper.Error($"退出登录请求失败：{ex.Message}");
                // 即使请求失败，也认为退出操作完成（因为本地即将清理）
                return AjaxResult.Error($"退出失败: {ex.Message}");
            }
        }
        
        /// <summary>
        /// 检查Token状态
        /// </summary>
        public static async Task<AjaxResult> CheckTokenAsync(string token, string deviceId)
        {
             /*
             var request = new HttpRequestMessage(HttpMethod.Post, $"{BaseApiUrl}/checkToken");
             AttachHeaders(request, token, deviceId);
             */
             await Task.Delay(200);
             return AjaxResult.Success("Token有效");
        }
        
        /// <summary>
        /// 静默刷新Token
        /// // 登录业务场景：Token静默刷新参数（对应第4条）
        /// </summary>
        public static async Task<AjaxResult> SilentRefreshAsync(string token)
        {
             /* 
             // Update endpoint to autoRefresh
             var request = new HttpRequestMessage(HttpMethod.Post, $"{BaseApiUrl}/autoRefresh"); // /prod-api/out/auth/autoRefresh
             AttachHeaders(request, token);
             */
             await Task.Delay(200);
             return AjaxResult.Success("刷新成功", (object?)Guid.NewGuid().ToString("N")); // Cast explicit to solve ambiguity with (msg,data) vs (data,msg)
        }
        
        /// <summary>
        /// 更新在线活跃状态
        /// // 登录业务场景：断网/联网状态同步逻辑（对应第7条）
        /// </summary>
        public static async Task<AjaxResult> UpdateActiveStatusAsync(long loginRecordId, string status)
        {
            /*
            var payload = new { loginRecordId = loginRecordId, loginStatus = status };
            // ...
            */
            await Task.Delay(100);
            return AjaxResult.Success("状态已更新");
        }

        /// <summary>
        /// 获取最新版本信息
        /// </summary>
        public static async Task<AjaxResult> GetLatestVersionAsync()
        {
             if (!NetworkHelper.IsNetworkAvailable()) return AjaxResult.Error("无网络连接");

             try
             {
                // 发起请求 (假设版本接口为 /out/common/client/version)
                // 若后端暂无此接口，可暂时保持 silent fail 或根据实际情况调整
                var request = new HttpRequestMessage(HttpMethod.Get, $"{AppHost}/out/common/version/latest");
                 // 不强制需要Token，但也可能需要
                 // var user = LocalStorageHelper.GetUserInfo();
                 // if (user != null) AttachHeaders(request, user.Token);

                 var response = await SendRequestWithRetryAsync(request, false);
                 if (response.IsSuccessStatusCode)
                 {
                     var result = await response.Content.ReadFromJsonAsync<AjaxResult>();
                     if (result != null && result.IsSuccess && result.Data is JsonElement je)
                     {
                         var v = new ViewModels.VersionInfo();
                         if (je.TryGetProperty("version", out var pv)) v.Version = pv.GetString() ?? "";
                         if (je.TryGetProperty("updateContent", out var pc)) v.UpdateContent = pc.GetString() ?? "";
                         return AjaxResult.Success("获取成功", v);
                     }
                 }
                 return AjaxResult.Error("获取版本信息失败");
             }
             catch(Exception ex)
             {
                 return AjaxResult.Error($"版本检查异常: {ex.Message}");
             }
        }

        /// <summary>
        /// 创建充值订单
        /// </summary>
        public static async Task<AjaxResult> CreateRechargeOrderAsync(long userId, string userName, decimal amount)
        {
            if (!NetworkHelper.IsNetworkAvailable())
            {
                return AjaxResult.Error("无网络连接，无法发起充值");
            }
            
            var user = LocalStorageHelper.GetUserInfo();
            // 登录业务场景：获取当前登录用户的Token (clientToken)
            // 注意：此处 clientToken 即为 Authorization Header 中使用的 Bearer Token
            string clientToken = user?.Token ?? string.Empty;

            try
            {
                var payload = new 
                { 
                    userId = userId, 
                    userName = userName, 
                    amount = amount, 
                    clientToken = clientToken // 用户登录凭证
                };

                var request = new HttpRequestMessage(HttpMethod.Post, $"{AppHost}/out/pay/order/initiate");
                request.Content = JsonContent.Create(payload);
                AttachHeaders(request, clientToken);

                var response = await SendRequestWithRetryAsync(request);
                var result = await response.Content.ReadFromJsonAsync<AjaxResult>();
                return result ?? AjaxResult.Error("服务器未返回有效数据");
            }
            catch (Exception ex)
            {
                LogHelper.Error($"创建充值订单失败：{ex.Message}");
                return AjaxResult.Error($"创建订单请求失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 查询充值订单状态
        /// </summary>
        public static async Task<AjaxResult> QueryRechargeOrderStatusAsync(string orderNo)
        {
            if (!NetworkHelper.IsNetworkAvailable()) return AjaxResult.Error("无网络连接");

            try
            {
                // 轮询请求携带订单号
                var payload = new { orderNo = orderNo };
                var request = new HttpRequestMessage(HttpMethod.Post, $"{AppHost}/out/pay/order/query");
                request.Content = JsonContent.Create(payload);
                 
                // 携带Token
                var user = LocalStorageHelper.GetUserInfo();
                if (user != null) AttachHeaders(request, user.Token);

                var response = await SendRequestWithRetryAsync(request);
                var result = await response.Content.ReadFromJsonAsync<AjaxResult>();
                
                return result ?? AjaxResult.Error("查询失败");
            }
            catch (Exception ex)
            {
                 // Polling errors are often ignored or logged quietly
                 // LogHelper.Debug($"查询订单状态异常: {ex.Message}");
                 return AjaxResult.Error("查询异常");
            }
        }
        
        /// <summary>
        /// 同步消费记录
        /// </summary>
        public static async Task<AjaxResult> SyncConsumeRecordsAsync(List<ConsumeRecord> records)
        {
            await Task.Delay(1000);

            if (records == null || records.Count == 0)
            {
                return AjaxResult.Success("无待同步的消费记录");
            }

            if (!NetworkHelper.IsNetworkAvailable())
            {
                return AjaxResult.Error("无网络连接，无法同步消费记录");
            }

            // 模拟同步成功，标记所有记录为已同步
            foreach (var record in records)
            {
                LocalDbHelper.MarkRecordAsSynced(record.RecordId);
            }

            LogHelper.Info($"成功同步{records.Count}条消费记录到后台");
            return AjaxResult.Success($"同步成功，共{records.Count}条记录");
        }

        /// <summary>
        /// 同步用户额度接口
        /// </summary>
        public static async Task<AjaxResult> SyncQuotaAsync(long userId, string username)
        {
            if (!NetworkHelper.IsNetworkAvailable())
            {
                return AjaxResult.Error("无网络连接，无法同步额度");
            }

            try
            {
                var payload = new { userId = userId, username = username };
                var request = new HttpRequestMessage(HttpMethod.Post, $"{AppHost}/out/common/user/syncQuota");
                request.Content = JsonContent.Create(payload);
                
                var user = LocalStorageHelper.GetUserInfo();
                if (user != null) AttachHeaders(request, user.Token);

                var response = await SendRequestWithRetryAsync(request);
                var result = await response.Content.ReadFromJsonAsync<AjaxResult>();
                return result ?? AjaxResult.Error("服务器未返回有效数据");
            }
            catch (Exception ex)
            {
                LogHelper.Error($"同步额度请求失败：{ex.Message}");
                return AjaxResult.Error($"同步额度失败: {ex.Message}");
            }
        }


        // 后端接口地址（可配置到AppConfig.json中，此处为示例）
        private const string _uploadLogApi = "https://your-backend-api.com/api/Log/Upload";

        /// <summary>
        /// 上传日志内容到后端（异步方法，避免阻塞UI）
        /// </summary>
        /// <param name="userAccount">用户账号（关联日志归属）</param>
        /// <param name="logFileName">日志文件名</param>
        /// <param name="logContent">日志文件完整内容</param>
        /// <returns>上传结果（成功/失败）</returns>
        public static async Task<bool> UploadLogToBackendAsync(string userAccount, string logFileName, string logContent)
        {
            if (string.IsNullOrEmpty(userAccount) || string.IsNullOrEmpty(logFileName) || string.IsNullOrEmpty(logContent))
            {
                LogHelper.Warn("上传日志失败：必要参数为空");
                return false;
            }

            // 模拟上传成功
            await Task.Delay(500);
            LogHelper.Info("日志上传成功（模拟）");
            return true;
        }

    }
}
