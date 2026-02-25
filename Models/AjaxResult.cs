using System;
using System.Collections.Generic;
using System.Text;
using Newtonsoft.Json;

namespace docment_tools_client.Models
{
    /// <summary>
    /// 适配RuoYi框架的接口返回结果模型
    /// </summary>
    public class AjaxResult
    {
        /// <summary>
        /// 状态码（200成功，500失败，401未授权）
        /// </summary>
        /// [JsonProperty("code")]
        public int Code { get; set; }

        public bool IsSuccess => Code == 200;

        /// <summary>
        /// 返回消息
        /// </summary>
        /// [JsonProperty("msg")]
        public string Msg { get; set; } = string.Empty;

        /// <summary>
        /// 返回数据
        /// </summary>
        /// [JsonProperty("data")]
        public object? Data { get; set; }

        /// <summary>
        /// 构建成功返回结果
        /// </summary>
        /// <param name="data">返回数据</param>
        /// <param name="msg">成功消息</param>
        /// <returns>AjaxResult</returns>
        public static AjaxResult Success(object? data = null, string msg = "操作成功")
        {
            return new AjaxResult
            {
                Code = 200,
                Msg = msg,
                Data = data
            };
        }

        /// <summary>
        /// 兼容调用：Success(string msg, object data)
        /// </summary>
        public static AjaxResult Success(string msg, object? data)
        {
            return new AjaxResult
            {
                Code = 200,
                Msg = msg,
                Data = data
            };
        }

        /// <summary>
        /// 构建失败返回结果
        /// </summary>
        /// <param name="msg">失败消息</param>
        /// <param name="code">错误状态码</param>
        /// <returns>AjaxResult</returns>
        public static AjaxResult Error(string msg = "操作失败", int code = 500)
        {
            return new AjaxResult
            {
                Code = code,
                Msg = msg,
                Data = null
            };
        }

        /// <summary>
        /// 构建未授权返回结果
        /// </summary>
        public static AjaxResult Unauthorized(string msg = "未登录或登录已过期")
        {
            return new AjaxResult
            {
                Code = 401,
                Msg = msg,
                Data = null
            };
        }
    }
}

