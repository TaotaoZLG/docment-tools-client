using docment_tools_client.Helpers;
using docment_tools_client.Models;
using MiniExcelLibs;
using NPOI.HSSF.UserModel; // For legacy xls
using NPOI.SS.UserModel;   // For IWorkbook, ISheet etc
using System.IO;
using System.Text;

namespace docment_tools_client.Services
{
    /// <summary>
    /// Excel读取/导出服务（基于MiniExcel + NPOI，支持.NET 10）
    /// 支持 xlsx (MiniExcel) 和 xls (NPOI)
    /// </summary>
    public static class ExcelService
    {
        /// <summary>
        /// 数据分割核心方法
        /// 根据填充人数将案件数据分割为多个子数组
        /// </summary>
        /// <param name="allData">完整案件数据列表</param>
        /// <param name="fillCount">填充人数限制（0或空表示不分割）</param>
        /// <returns>分割后的二维数组</returns>
        public static List<List<DynamicCaseData>> SplitCaseData(List<DynamicCaseData> allData, int fillCount)
        {
            var result = new List<List<DynamicCaseData>>();
            if (allData == null || allData.Count == 0) return result;

            // 1. 如果填充人数为空或0，或者大于等于总条数，则不分割，返回整体
            if (fillCount <= 0 || fillCount >= allData.Count)
            {
                result.Add(new List<DynamicCaseData>(allData));
                return result;
            }

            // 2. 按数值分割
            for (int i = 0; i < allData.Count; i += fillCount)
            {
                // Math.Min 防止越界
                int count = Math.Min(fillCount, allData.Count - i);
                var chunk = allData.GetRange(i, count);
                result.Add(chunk);
            }

            LogHelper.Info($"数据分割完成：总条数{allData.Count}，每份{fillCount}条，共生成{result.Count}份文档");
            return result;
        }

        /// <summary>
        /// 读取Excel文件，返回案件数据列表和表头列表
        /// </summary>
        /// <param name="excelPath">Excel文件路径</param>
        /// <returns>表头列表 + 案件数据列表</returns>
        public static (List<string> Headers, List<DynamicCaseData> DynamicCaseDatas) ReadExcel(string excelPath)
        {
            // 根据扩展名检查是否为 .xls
            var extension = Path.GetExtension(excelPath).ToLower();
            if (extension == ".xls")
            {
                return ReadXlsWithNpoi(excelPath);
            }

            var headers = new List<string>();
            var dynamicCaseDatas = new List<DynamicCaseData>();

            try
            {
                if (!File.Exists(excelPath))
                {
                    LogHelper.Error($"Excel文件不存在：{excelPath}");
                    return (headers, dynamicCaseDatas);
                }

                // 使用MiniExcel读取所有行，强制使用第一行作为表头
                // 使用FileStream以支持打开的文件读取
                using var stream = new FileStream(excelPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                var rows = stream.Query(useHeaderRow: true).ToList();

                if (rows.Count == 0)

                {
                    LogHelper.Warn("Excel文件中无数据");
                    return (headers, dynamicCaseDatas);
                }

                // 1. 读取表头
                var firstRow = rows.FirstOrDefault() as IDictionary<string, object>;
                if (firstRow != null)
                {
                    headers = firstRow.Keys.Select(k => k.Trim()).ToList();
                }

                if (headers.Count == 0)
                {
                    LogHelper.Warn("未能识别Excel表头，请检查第一行是否为空");
                }

                // 2. 转换数据
                foreach (IDictionary<string, object> row in rows)
                {
                    var dynamicCaseData = new DynamicCaseData();
                    bool isRowEmpty = true;

                    foreach (var key in row.Keys)
                    {
                        var cleanKey = key.Trim();
                        // 增强类型转换逻辑
                        var value = ConvertCellValueToString(row[key], cleanKey);
                        
                        if (!string.IsNullOrWhiteSpace(value))
                        {
                            isRowEmpty = false;
                        }

                        if (!dynamicCaseData.CaseInfo.ContainsKey(cleanKey))
                        {
                            dynamicCaseData.CaseInfo.Add(cleanKey, value);
                        }
                    }
                    
                    if (!isRowEmpty)
                    {
                        dynamicCaseDatas.Add(dynamicCaseData);
                    }
                }


                if (dynamicCaseDatas.Count > 0)
                {
                    var sample = dynamicCaseDatas[0];
                    var sb = new StringBuilder();
                    foreach(var kv in sample.CaseInfo) sb.Append($"{kv.Key}={kv.Value}; ");
                    LogHelper.Info($"[MiniExcel] 数据样例(第1条): {sb.ToString()}");
                }

                LogHelper.Info($"成功读取Excel文件（xlsx/csv），共{dynamicCaseDatas.Count}条案件数据");
            }
            catch (Exception ex)
            {
                LogHelper.Error($"读取Excel文件失败：{ex.Message}");
            }

            return (headers, dynamicCaseDatas);
        }

        /// <summary>
        /// 使用 NPOI 读取旧版 Excel (.xls)
        /// </summary>
        private static (List<string> Headers, List<DynamicCaseData> DynamicCaseDatas) ReadXlsWithNpoi(string excelPath)
        {
            var headers = new List<string>();
            var dynamicCaseDatas = new List<DynamicCaseData>();

            try
            {
                if (!File.Exists(excelPath))
                {
                    LogHelper.Error($"Excel文件不存在：{excelPath}");
                    return (headers, dynamicCaseDatas);
                }

                using var fileStream = new FileStream(excelPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                var workbook = new HSSFWorkbook(fileStream); // HSSF for .xls
                var sheet = workbook.GetSheetAt(0); // 读取第一个Sheet


                if (sheet == null || sheet.LastRowNum < 0)
                {
                    LogHelper.Warn("[NPOI] Excel Sheet为空");
                    return (headers, dynamicCaseDatas);
                }

                // 1. 读取表头（第一行）
                var headerRow = sheet.GetRow(0);
                if (headerRow == null)
                {
                    LogHelper.Warn("[NPOI] 未找到表头行");
                    return (headers, dynamicCaseDatas);
                }

                for (int i = 0; i < headerRow.LastCellNum; i++)
                {
                    var cell = headerRow.GetCell(i);
                    headers.Add(cell?.ToString()?.Trim() ?? $"Column{i}");
                }

                // 2. 读取数据（从第二行开始）
                for (int i = 1; i <= sheet.LastRowNum; i++)
                {
                    var row = sheet.GetRow(i);
                    if (row == null) continue;

                    var dynamicCaseData = new DynamicCaseData();
                    // 确保不越界，且每一列都读取（即使该行某些列为空）
                    int cellCount = Math.Max(headers.Count, row.LastCellNum);
                    bool isRowEmpty = true;

                    for (int j = 0; j < headers.Count; j++)
                    {
                        var key = headers[j];
                        var cell = row.GetCell(j);
                        string value = string.Empty;

                        if (cell != null)
                        {
                            // NPOI 单元格类型处理
                            switch (cell.CellType)
                            {
                                case CellType.Numeric:
                                    if (DateUtil.IsCellDateFormatted(cell))
                                    {
                                        var dt = cell.DateCellValue;
                                        value = dt == null ? string.Empty : ((DateTime)dt).ToString("yyyy 年 MM 月 dd 日");
                                    }
                                    else
                                    {
                                        value = cell.NumericCellValue.ToString();
                                    }
                                    break;
                                case CellType.Boolean:
                                    value = cell.BooleanCellValue.ToString();
                                    break;
                                case CellType.Formula:
                                    try { value = cell.StringCellValue; }
                                    catch { value = cell.NumericCellValue.ToString(); }
                                    break;
                                default:
                                    value = cell.ToString()?.Trim() ?? string.Empty;
                                    break;
                            }
                        }

                        if (!string.IsNullOrWhiteSpace(value))
                        {
                            isRowEmpty = false;
                        }

                        if (!dynamicCaseData.CaseInfo.ContainsKey(key))
                        {
                            dynamicCaseData.CaseInfo.Add(key, value);
                        }
                    }
                    
                    // 只有当行内有实际数据时才添加
                    if (!isRowEmpty)
                    {
                        dynamicCaseDatas.Add(dynamicCaseData);
                    }

                }

                if (dynamicCaseDatas.Count > 0)
                {
                    var sample = dynamicCaseDatas[0];
                    var sb = new StringBuilder();
                    foreach(var kv in sample.CaseInfo) sb.Append($"{kv.Key}={kv.Value}; ");
                    LogHelper.Info($"[NPOI] 数据样例(第1条): {sb.ToString()}");
                }

                LogHelper.Info($"成功读取Excel (.xls) 文件，共{dynamicCaseDatas.Count}条案件数据");
            }
            catch (Exception ex)
            {
                LogHelper.Error($"[NPOI] 读取Excel文件失败：{ex.Message}");
            }

            return (headers, dynamicCaseDatas);
        }


        /// <summary>
        /// 灵活转换单元格值为字符串（处理日期、数字等格式）
        /// </summary>
        private static string ConvertCellValueToString(object? cellValue, string columnName = "")
        {
            if (cellValue == null) return string.Empty;

            try 
            {
                // 1. 处理日期类型
                if (cellValue is DateTime dt)
                {
                    // 格式："yyyy 年 MM 月 dd 日" (数字和文字之间包含空格)
                    return dt.TimeOfDay == TimeSpan.Zero 
                        ? dt.ToString("yyyy 年 MM 月 dd 日") 
                        : dt.ToString("yyyy 年 MM 月 dd 日 HH:mm:ss");
                }
                
                // 2. 处理字节数组 (防止二进制数据ToString)
                if (cellValue is byte[])
                {
                     return "(System.Byte[])";
                }

                // 3. 处理数值型日期 (如 46033 -> 2026/1/11)
                // 只有当列名包含日期相关关键字，且数值在合理日期范围内时才转换
                if (cellValue is double || cellValue is int || cellValue is decimal)
                {
                     double dVal = Convert.ToDouble(cellValue);
                     // 简单判断列名是否包含 "日期", "Date", "时间", "Time"
                     bool looksLikeDateColumn = 
                         columnName.Contains("日期") || 
                         columnName.Contains("Date", StringComparison.OrdinalIgnoreCase) ||
                         columnName.Contains("时间") ||
                         columnName.Contains("Time", StringComparison.OrdinalIgnoreCase);

                     // OADate 范围：10959 (1930年) ~ 73050 (2100年)，避免误伤金额等普通数字
                     if (looksLikeDateColumn && dVal > 10959 && dVal < 73050)
                     {
                         try 
                         {
                             var dateFromOA = DateTime.FromOADate(dVal);
                             return dateFromOA.TimeOfDay == TimeSpan.Zero 
                                ? dateFromOA.ToString("yyyy 年 MM 月 dd 日") 
                                : dateFromOA.ToString("yyyy 年 MM 月 dd 日 HH:mm:ss");
                         }
                         catch { }
                     }
                }

                // 4. 通用处理
                return cellValue.ToString()?.Trim() ?? string.Empty;
            }
            catch
            {
                return cellValue?.ToString() ?? string.Empty;
            }
        }
       
        /// <summary>
        /// 导出案件数据模板（空白Excel，包含默认表头）
        /// </summary>
        /// <param name="targetPath">导出目标路径</param>
        /// <param name="defaultHeaders">默认表头（可选）</param>
        public static bool ExportCaseTemplate(string targetPath, List<string>? defaultHeaders = null)
        {
            try
            {
                // 初始化默认表头
                var headers = defaultHeaders ?? new List<string> { "姓名", "身份证号", "案由", "案件金额（元）", "案件日期" };

                // 创建一个包含表头的字典列表（仅一行空数据，或者仅表头）
                // MiniExcel SaveAs takes IEnumerable. To just creating headers, we can create a List<Dictionary<string, object>>.
                
                var data = new List<Dictionary<string, object>>();
                // Create an empty row structure
                // But MiniExcel needs at least one object to know column names if we rely on implicit mapping,
                // OR we just save an empty enumerable but that might not create headers.
                
                // Better approach: Create a dictionary for the first row with empty values to establish headers.
                var templateRow = new Dictionary<string, object>();
                foreach (var header in headers)
                {
                    templateRow[header] = "";
                }
                data.Add(templateRow);

                // Save to file. MiniExcel will write headers based on dictionary keys.
                // Overwrite property is default true for SaveAs? No, MiniExcel creates new file.
                // Use FileStream to be safe or just path string.
                
                MiniExcel.SaveAs(targetPath, data);

                LogHelper.Info($"成功导出案件数据模板：{targetPath}");
                return true;
            }
            catch (Exception ex)
            {
                LogHelper.Error($"导出案件数据模板失败：{ex.Message}");
                return false;
            }
        }
    }
}
