using docment_tools_client.Models;
using Microsoft.Data.Sqlite;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace docment_tools_client.Helpers
{
    /// <summary>
    /// SQLite本地数据库工具类（真实SQLite操作，持久化存储消费记录）
    /// 用于存储消费记录、未同步数据
    /// </summary>
    public static class LocalDbHelper
    {
        /// <summary>
        /// SQLite数据库文件路径（%AppData%\Local\DocumentToolsClient\DocumentToolsClient.db）
        /// </summary>
        private static readonly string _sqliteDbPath;

        /// <summary>
        /// 数据库连接字符串
        /// </summary>
        private static readonly string _connectionString;

        /// <summary>
        /// 静态构造函数（初始化数据库路径、连接字符串，自动创建数据库和表，自动清理过期已同步记录）
        /// </summary>
        static LocalDbHelper()
        {
            // 1. 初始化数据库文件路径
            var appDataDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "DocumentToolsClient");

            if (!Directory.Exists(appDataDir))
            {
                Directory.CreateDirectory(appDataDir);
            }

            _sqliteDbPath = Path.Combine(appDataDir, "DocumentToolsClient.db");
            _connectionString = $"Data Source={_sqliteDbPath};Cache=Shared;";

            // 2. 自动初始化数据库（创建消费记录表，若不存在）
            InitSqliteDatabase();

            // 3. 自动删除一周前已同步的消费记录（应用启动时执行一次，可根据需求调整执行时机）
            DeleteExpiredSyncedRecords(7); // 传入7天，代表删除7天前（一周前）的已同步记录
        }

        /// <summary>
        /// 初始化SQLite数据库（创建必要的表）
        /// </summary>
        private static void InitSqliteDatabase()
        {
            try
            {
                // 连接SQLite数据库（不存在则自动创建.db文件）
                using var connection = new SqliteConnection(_connectionString);
                connection.Open();

                // 3. 创建消费记录表（ConsumeRecords）
                // 字段对应ConsumeRecord模型，SQLite字段类型适配：
                // - 字符串：TEXT
                // - 整数：INTEGER
                // - 小数：REAL（适配decimal类型）
                // - 日期时间：TEXT（存储为ISO标准格式，便于序列化/反序列化）
                // - 布尔值：INTEGER（0=false，1=true）
                var createTableSql = @"
                    CREATE TABLE IF NOT EXISTS ConsumeRecords (
                        RecordId TEXT PRIMARY KEY NOT NULL,
                        UserAccount TEXT NOT NULL,
                        CaseCount INTEGER NOT NULL DEFAULT 0,
                        Amount REAL NOT NULL DEFAULT 0.0,
                        ConsumeTime TEXT NOT NULL,
                        IsSynced INTEGER NOT NULL DEFAULT 0,
                        DocumentType TEXT DEFAULT '',
                        DocumentCaseCount INTEGER NOT NULL DEFAULT 0,
                        FilePath TEXT DEFAULT ''
                    );";

                using var command = new SqliteCommand(createTableSql, connection);
                command.ExecuteNonQuery();

                // 尝试添加新列 FilePath（用于兼容旧版本数据库）
                try
                {
                    var alterSql = "ALTER TABLE ConsumeRecords ADD COLUMN FilePath TEXT DEFAULT '';";

                    using var alterCmd = new SqliteCommand(alterSql, connection);
                    alterCmd.ExecuteNonQuery();
                }
                catch
                {
                    // 忽略错误（说明列已存在）
                }

                LogHelper.Info($"SQLite数据库初始化成功，数据库文件路径：{_sqliteDbPath}");
            }
            catch (Exception ex)
            {
                LogHelper.Error($"SQLite数据库初始化失败：{ex.Message}");
            }
        }

        /// <summary>
        /// 添加消费记录到SQLite数据库（持久化存储）
        /// </summary>
        public static void AddConsumeRecord(ConsumeRecord record)
        {
            if (record == null)
            {
                LogHelper.Warn("添加消费记录失败：传入的ConsumeRecord对象为null");
                return;
            }

            try
            {
                using var connection = new SqliteConnection(_connectionString);
                connection.Open();

                // 插入SQL（参数化查询，防止SQL注入）
                var insertSql = @"
                    INSERT OR REPLACE INTO ConsumeRecords (
                        RecordId, UserAccount, CaseCount, Amount, 
                        ConsumeTime, IsSynced, DocumentType, DocumentCaseCount, FilePath
                    ) VALUES (
                        @RecordId, @UserAccount, @CaseCount, @Amount,
                        @ConsumeTime, @IsSynced, @DocumentType, @DocumentCaseCount, @FilePath
                    );";

                using var command = new SqliteCommand(insertSql, connection);

                // 绑定参数（映射ConsumeRecord模型字段，处理数据类型转换）
                command.Parameters.AddWithValue("@RecordId", record.RecordId);
                command.Parameters.AddWithValue("@UserAccount", record.UserAccount);
                command.Parameters.AddWithValue("@CaseCount", record.CaseCount);
                command.Parameters.AddWithValue("@Amount", (double)record.Amount); // decimal转double适配SQLite REAL
                command.Parameters.AddWithValue("@ConsumeTime", record.ConsumeTime.ToString("o")); // ISO标准日期格式
                command.Parameters.AddWithValue("@IsSynced", record.IsSynced ? 1 : 0); // 布尔值转整数
                command.Parameters.AddWithValue("@DocumentType", record.DocumentType);
                command.Parameters.AddWithValue("@DocumentCaseCount", record.DocumentCaseCount);
                command.Parameters.AddWithValue("@FilePath", record.FilePath ?? string.Empty);

                // 执行插入操作
                command.ExecuteNonQuery();

                LogHelper.Info($"成功添加消费记录到SQLite数据库，记录ID：{record.RecordId}");
            }
            catch (Exception ex)
            {
                LogHelper.Error($"添加消费记录到SQLite数据库失败：{ex.Message}，记录ID：{record.RecordId}");
            }
        }

        /// <summary>
        /// 从SQLite数据库获取本地未同步的消费记录
        /// </summary>
        public static List<ConsumeRecord> GetUnsyncedConsumeRecords(string userAccount)
        {
            var unsyncedRecords = new List<ConsumeRecord>();

            if (string.IsNullOrEmpty(userAccount))
            {
                LogHelper.Warn("获取未同步消费记录失败：用户账号为空");
                return unsyncedRecords;
            }

            try
            {
                using var connection = new SqliteConnection(_connectionString);
                connection.Open();

                // 查询SQL：筛选IsSynced=0（未同步）且指定用户账号的记录
                var selectSql = @"
                    SELECT * FROM ConsumeRecords 
                    WHERE UserAccount = @UserAccount AND IsSynced = 0
                    ORDER BY ConsumeTime DESC;";

                using var command = new SqliteCommand(selectSql, connection);
                command.Parameters.AddWithValue("@UserAccount", userAccount);

                // 执行查询并读取结果
                using var reader = command.ExecuteReader();
                while (reader.Read())
                {
                    var record = MapReaderToConsumeRecord(reader);
                    if (record != null)
                    {
                        unsyncedRecords.Add(record);
                    }
                }

                LogHelper.Info($"成功从SQLite数据库获取{unsyncedRecords.Count}条未同步消费记录，用户账号：{userAccount}");
            }
            catch (Exception ex)
            {
                LogHelper.Error($"获取未同步消费记录失败：{ex.Message}，用户账号：{userAccount}");
            }

            return unsyncedRecords;
        }

        /// <summary>
        /// 标记消费记录为已同步（更新SQLite数据库中的IsSynced字段）
        /// </summary>
        public static void MarkRecordAsSynced(string recordId)
        {
            if (string.IsNullOrEmpty(recordId))
            {
                LogHelper.Warn("标记消费记录为已同步失败：记录ID为空");
                return;
            }

            try
            {
                using var connection = new SqliteConnection(_connectionString);
                connection.Open();

                // 更新SQL：将IsSynced设为1（已同步）
                var updateSql = @"
                    UPDATE ConsumeRecords 
                    SET IsSynced = 1 
                    WHERE RecordId = @RecordId;";

                using var command = new SqliteCommand(updateSql, connection);
                command.Parameters.AddWithValue("@RecordId", recordId);

                // 执行更新操作
                var affectedRows = command.ExecuteNonQuery();
                if (affectedRows > 0)
                {
                    LogHelper.Info($"成功标记消费记录为已同步，记录ID：{recordId}");
                }
                else
                {
                    LogHelper.Warn($"未找到指定的消费记录，无法标记为已同步，记录ID：{recordId}");
                }
            }
            catch (Exception ex)
            {
                LogHelper.Error($"标记消费记录为已同步失败：{ex.Message}，记录ID：{recordId}");
            }
        }

        /// <summary>
        /// 从SQLite数据库获取指定用户的所有消费记录
        /// </summary>
        public static List<ConsumeRecord> GetAllConsumeRecords(string userAccount)
        {
            var allRecords = new List<ConsumeRecord>();

            if (string.IsNullOrEmpty(userAccount))
            {
                LogHelper.Warn("获取所有消费记录失败：用户账号为空");
                return allRecords;
            }

            try
            {
                using var connection = new SqliteConnection(_connectionString);
                connection.Open();

                // 查询SQL：获取指定用户的所有消费记录，按消费时间倒序
                var selectSql = @"
                    SELECT * FROM ConsumeRecords 
                    WHERE UserAccount = @UserAccount
                    ORDER BY ConsumeTime DESC;";

                using var command = new SqliteCommand(selectSql, connection);
                command.Parameters.AddWithValue("@UserAccount", userAccount);

                // 执行查询并读取结果
                using var reader = command.ExecuteReader();
                while (reader.Read())
                {
                    var record = MapReaderToConsumeRecord(reader);
                    if (record != null)
                    {
                        allRecords.Add(record);
                    }
                }

                LogHelper.Info($"成功从SQLite数据库获取{allRecords.Count}条消费记录，用户账号：{userAccount}");
            }
            catch (Exception ex)
            {
                LogHelper.Error($"获取所有消费记录失败：{ex.Message}，用户账号：{userAccount}");
            }

            return allRecords;
        }

        /// <summary>
        /// 删除指定天数前、已同步的消费记录（核心清理方法）
        /// </summary>
        /// <param name="expireDays">过期天数，默认7天（一周）</param>
        public static void DeleteExpiredSyncedRecords(int expireDays = 7)
        {
            try
            {
                using var connection = new SqliteConnection(_connectionString);
                connection.Open();

                // 删除SQL：双重筛选（已同步 + 消费时间早于指定天数前）
                // 利用SQLite的datetime函数处理ISO格式日期字符串，无需C#转换，更准确
                var deleteSql = @"
                    DELETE FROM ConsumeRecords 
                    WHERE IsSynced = 1 
                    AND datetime(ConsumeTime) < datetime('now', '-' || @ExpireDays || ' days');";

                using var command = new SqliteCommand(deleteSql, connection);
                command.Parameters.AddWithValue("@ExpireDays", expireDays); // 传入过期天数，灵活配置

                // 执行删除操作，返回受影响的行数（即删除的记录数）
                var deletedCount = command.ExecuteNonQuery();

                LogHelper.Info($"成功清理{deletedCount}条{expireDays}天前已同步的消费记录");
            }
            catch (Exception ex)
            {
                LogHelper.Error($"清理过期已同步消费记录失败：{ex.Message}");
            }
        }

        /// <summary>
        /// 辅助方法：将SqliteDataReader映射为ConsumeRecord对象（处理数据类型转换）
        /// </summary>
        private static ConsumeRecord? MapReaderToConsumeRecord(SqliteDataReader reader)
        {
            try
            {
                return new ConsumeRecord
                {
                    RecordId = reader["RecordId"].ToString() ?? string.Empty,
                    UserAccount = reader["UserAccount"].ToString() ?? string.Empty,
                    CaseCount = Convert.ToInt32(reader["CaseCount"]),
                    Amount = Convert.ToDecimal(reader["Amount"]), // double转decimal恢复额度
                    ConsumeTime = DateTime.Parse(reader["ConsumeTime"].ToString() ?? DateTime.Now.ToString("o")),
                    IsSynced = Convert.ToInt32(reader["IsSynced"]) == 1, // 整数转布尔值
                    DocumentType = reader["DocumentType"].ToString() ?? string.Empty,
                    DocumentCaseCount = Convert.ToInt32(reader["DocumentCaseCount"]),
                    FilePath = reader.GetOrdinal("FilePath") >= 0 && !reader.IsDBNull(reader.GetOrdinal("FilePath")) 
                                ? reader["FilePath"].ToString() ?? string.Empty 
                                : string.Empty
                };
            }
            catch (Exception ex)
            {
                LogHelper.Error($"映射ConsumeRecord对象失败：{ex.Message}");
                return null;
            }
        }
    }
}
