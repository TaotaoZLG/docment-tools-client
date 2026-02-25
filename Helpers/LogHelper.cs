using Serilog;
using Serilog.Events;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Windows;

namespace docment_tools_client.Helpers
{
    public enum LogLevel
    {
        Info,
        Warn,
        Error
    }

    /// <summary>
    /// 日志条目模型（用于UI绑定展示，支持属性变更通知）
    /// </summary>
    public class LogItem : INotifyPropertyChanged
    {
        private string _logLevel = string.Empty;
        private string _content = string.Empty;
        private DateTime _createTime;

        public string LogLevel
        {
            get => _logLevel;
            set { _logLevel = value; OnPropertyChanged(nameof(LogLevel)); }
        }

        public string Content
        {
            get => _content;
            set { _content = value; OnPropertyChanged(nameof(Content)); }
        }

        public DateTime CreateTime
        {
            get => _createTime;
            set { _createTime = value; OnPropertyChanged(nameof(CreateTime)); }
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    /// <summary>
    /// 日志工具类
    /// 核心功能：1. UI实时展示（ObservableCollection绑定） 2. Serilog本地按日存储 3. 自动删除7天前过期日志 4. 支持日志上传后端前置准备
    /// </summary>
    public static class LogHelper
    {
        public static event Action<string, LogLevel>? OnLogReceived;

        /// <summary>
        /// 实时日志集合（用于UI绑定，ObservableCollection支持自动刷新UI，无需手动通知）
        /// </summary>
        public static readonly ObservableCollection<LogItem> LiveLogCollection = new ObservableCollection<LogItem>();

        /// <summary>
        /// 日志文件存储目录
        /// </summary>
        private static readonly string _logDir;

        /// <summary>
        /// 日志文件前缀（匹配Serilog生成的文件名格式）
        /// </summary>
        private static readonly string _logFilePrefix = "DocumentToolsClient_";

        /// <summary>
        /// 过期日志天数（默认7天）
        /// </summary>
        private const int _expireDays = 7;

        /// <summary>
        /// 静态构造函数（初始化目录、Serilog、自动清理过期日志）
        /// </summary>
        static LogHelper()
        {
            // 1. 初始化日志目录
            _logDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "DocumentToolsClient",
                "Logs");
            Directory.CreateDirectory(_logDir);

            // 2. 初始化Serilog（本地文件存储，按日期分割，生成格式：DocumentToolsClient_20260201.log）
            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Information()
                .WriteTo.File(
                    Path.Combine(_logDir, $"{_logFilePrefix}.log"),
                    rollingInterval: RollingInterval.Day, // 按日分割日志
                    retainedFileCountLimit: null, // 关闭Serilog自带的保留限制，改用自定义逻辑清理
                    outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss.fff} [{Level:u3}] {Message:lj}{NewLine}{Exception}")
                .CreateLogger();

            // 3. 自动清理7天前的过期本地日志文件（应用启动时执行一次）
            DeleteExpiredLogFiles();

            // 4. 注册应用退出事件，释放Serilog资源
            Application.Current.Exit += (s, e) => Log.CloseAndFlush();
        }

        /// <summary>
        /// 添加Info级别日志（绿色，普通信息）
        /// </summary>
        public static void Info(string content)
        {
            if (string.IsNullOrEmpty(content)) return;
            WriteLog(LogEventLevel.Information, content);
        }

        /// <summary>
        /// 添加Warn级别日志（深黄色/橙色，警告信息）
        /// </summary>
        public static void Warn(string content)
        {
            if (string.IsNullOrEmpty(content)) return;
            WriteLog(LogEventLevel.Warning, content);
        }

        /// <summary>
        /// 添加Error级别日志（红色，错误信息）
        /// </summary>
        public static void Error(string content)
        {
            if (string.IsNullOrEmpty(content)) return;
            WriteLog(LogEventLevel.Error, content);
        }

        /// <summary>
        /// 重载：添加Error级别日志（包含异常详情）
        /// </summary>
        public static void Error(string content, Exception ex)
        {
            if (string.IsNullOrEmpty(content)) content = "未知错误";
            var fullContent = $"{content} | 异常详情：{ex.Message}{Environment.NewLine}堆栈跟踪：{ex.StackTrace}";
            WriteLog(LogEventLevel.Error, fullContent);
        }

        /// <summary>
        /// 核心：写入日志（同步到本地文件 + UI实时集合）
        /// </summary>
        private static void WriteLog(LogEventLevel level, string content)
        {
            try
            {
                // 1. 写入Serilog本地文件（按日分割存储）
                switch (level)
                {
                    case LogEventLevel.Information:
                        Log.Information(content);
                        OnLogReceived?.Invoke(content, LogLevel.Info);
                        break;
                    case LogEventLevel.Warning:
                        Log.Warning(content);
                        OnLogReceived?.Invoke(content, LogLevel.Warn);
                        break;
                    case LogEventLevel.Error:
                        Log.Error(content);
                        OnLogReceived?.Invoke(content, LogLevel.Error);
                        break;
                    default:
                        Log.Information(content);
                        OnLogReceived?.Invoke(content, LogLevel.Info);
                        break;
                }

                // 2. 添加到实时日志集合（切换到UI线程，避免跨线程访问异常）
                Application.Current.Dispatcher.Invoke(() =>
                {
                    LiveLogCollection.Add(new LogItem
                    {
                        LogLevel = level.ToString().ToUpper(),
                        Content = content,
                        CreateTime = DateTime.Now
                    });

                    // 限制内存日志数量，避免内存溢出（保留最新1000条）
                    while (LiveLogCollection.Count > 1000)
                    {
                        LiveLogCollection.RemoveAt(0);
                    }
                });
            }
            catch (Exception ex)
            {
                // 避免日志写入失败导致主程序崩溃，仅输出到控制台
                Console.WriteLine($"日志写入失败：{ex.Message}");
            }
        }

        /// <summary>
        /// 自动删除7天前的本地日志文件
        /// </summary>
        private static void DeleteExpiredLogFiles()
        {
            try
            {
                if (!Directory.Exists(_logDir)) return;

                // 1. 获取目录下所有Serilog生成的日志文件
                var logFiles = Directory.GetFiles(_logDir, $"{_logFilePrefix}*.log")
                                        .Where(file => Path.GetFileName(file).StartsWith(_logFilePrefix))
                                        .ToList();

                if (logFiles.Count == 0) return;

                // 2. 计算过期时间（7天前的当前时间）
                var expireDateTime = DateTime.Now.AddDays(-_expireDays);
                int deletedCount = 0;

                // 3. 遍历文件，判断是否过期并删除
                foreach (var file in logFiles)
                {
                    try
                    {
                        // 解析文件名中的日期（格式：DocumentToolsClient_20260201.log）
                        var fileName = Path.GetFileNameWithoutExtension(file);
                        var dateStr = fileName.Replace(_logFilePrefix, string.Empty);
                        if (!DateTime.TryParseExact(dateStr, "yyyyMMdd", null, System.Globalization.DateTimeStyles.None, out var fileDate))
                        {
                            continue; // 非按日分割的日志文件，跳过
                        }

                        // 4. 若文件日期早于过期时间，删除该日志文件
                        if (fileDate < expireDateTime.Date)
                        {
                            File.Delete(file);
                            deletedCount++;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"删除单个过期日志文件失败（文件：{file}）：{ex.Message}");
                    }
                }

                // 5. 记录清理结果（仅信息日志，不干扰UI核心功能）
                Info($"已自动清理{deletedCount}条{_expireDays}天前的过期日志文件");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"批量清理过期日志文件失败：{ex.Message}");
            }
        }

        // ---------------------- 日志上传后端的前置辅助方法 ----------------------
        /// <summary>
        /// 获取指定日期范围内的本地日志文件列表（用于上传选择）
        /// </summary>
        /// <param name="startDate">开始日期（包含）</param>
        /// <param name="endDate">结束日期（包含）</param>
        /// <returns>日志文件路径列表</returns>
        public static List<string> GetLogFilesByDateRange(DateTime startDate, DateTime endDate)
        {
            var logFiles = new List<string>();
            try
            {
                if (!Directory.Exists(_logDir)) return logFiles;

                // 筛选指定日期范围内的日志文件
                logFiles = Directory.GetFiles(_logDir, $"{_logFilePrefix}*.log")
                                    .Where(file =>
                                    {
                                        var fileName = Path.GetFileNameWithoutExtension(file);
                                        var dateStr = fileName.Replace(_logFilePrefix, string.Empty);
                                        if (DateTime.TryParseExact(dateStr, "yyyyMMdd", null, System.Globalization.DateTimeStyles.None, out var fileDate))
                                        {
                                            return fileDate >= startDate.Date && fileDate <= endDate.Date;
                                        }
                                        return false;
                                    })
                                    .OrderBy(file => file)
                                    .ToList();
            }
            catch (Exception ex)
            {
                Error($"获取指定日期范围日志文件失败：{ex.Message}");
            }
            return logFiles;
        }

        /// <summary>
        /// 读取单个日志文件的完整内容（用于上传后端）
        /// </summary>
        /// <param name="filePath">日志文件路径</param>
        /// <returns>日志文件文本内容</returns>
        public static string ReadLogFileContent(string filePath)
        {
            try
            {
                if (!File.Exists(filePath)) return string.Empty;
                // 以UTF-8编码读取日志文件（匹配Serilog的写入编码）
                return File.ReadAllText(filePath, System.Text.Encoding.UTF8);
            }
            catch (Exception ex)
            {
                Error($"读取日志文件内容失败（文件：{Path.GetFileName(filePath)}）：{ex.Message}");
                return string.Empty;
            }
        }
    }
}
