using docment_tools_client.Helpers;
using docment_tools_client.Models;
using NPOI.XWPF.UserModel;
using NPOI.OpenXmlFormats.Wordprocessing;
using Newtonsoft.Json;
using System.IO;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace docment_tools_client.Services
{
    /// <summary>
    /// 升级版Word模板操作工具类（基于NPOI，支持单值填充、表格循环、段落循环）
    /// </summary>
    public static class WordService
    {
        // 标记常量
        private const string LoopStartPattern = "LOOP_(\\w+)_START"; // {{LOOP_Item_START}}
        // private const string LoopEndPattern = "LOOP_(\\w+)_END";     // {{LOOP_Item_END}}

        /// <summary>
        /// 获取默认模板路径
        /// </summary>
        public static string GetDefaultTemplatePath()
        {
            if (Directory.Exists(SystemTemplateDir))
            {
                var files = Directory.GetFiles(SystemTemplateDir, "*.docx");
                if (files.Length > 0) return files[0];
            }
            return string.Empty;
        }

        /// <summary>
        /// 预览Word文档
        /// </summary>
        public static void PreviewWord(string path)
        {
            if (File.Exists(path))
            {
                try
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(path) { UseShellExecute = true });
                }
                catch (Exception ex)
                {
                    LogHelper.Error($"无法预览文档：{path}", ex);
                }
            }
        }

        #region 核心生成方法

        /// <summary>
        /// 生成单份文书
        /// </summary>
        // 为了兼容旧代码，保留原签名。
        public static bool GenerateWord(string templatePath, DynamicCaseData caseData, string outputFilePath, bool withAttachment, bool isReadOnly = false)
        {
             // 简单的适配器：将单一 CaseData 视为一个包含单条数据的列表上下文？
             // 不，"单份文书" 仍然是 "Single Context"。
             // 但如果 caseData.CaseInfo 中包含 "Items" 列表同样可以处理 Loop。
             // 这里主要逻辑不变，我们添加一个新的重载或者在内部判断。
             return GenerateWordInternal(templatePath, caseData, outputFilePath, withAttachment, isReadOnly);
        }

        /// <summary>
        /// (新) 生成文书，支持 List<DynamicCaseData> 作为上下文（用于填充人数分割后的批量数据）
        /// </summary>
        public static bool GenerateWord(string templatePath, List<DynamicCaseData> caseDataList, string outputFilePath, bool withAttachment, bool isReadOnly = false)
        {
             return GenerateWordInternal(templatePath, caseDataList, outputFilePath, withAttachment, isReadOnly);
        }

        private static bool GenerateWordInternal(string templatePath, object dataContext, string outputFilePath, bool withAttachment, bool isReadOnly)
        {
            if (dataContext == null || string.IsNullOrEmpty(templatePath) || !File.Exists(templatePath))
            {
                LogHelper.Error($"生成失败：无效的参数或模板不存在 {templatePath}");
                return false;
            }

            try
            {
                // outputFilePath expects a full file path (e.g. C:\Docs\Contract_001.docx)
                string dir = Path.GetDirectoryName(outputFilePath);
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir)) 
                    Directory.CreateDirectory(dir);

                byte[] templateBytes;
                using (var fs = new FileStream(templatePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    using (var msTemp = new MemoryStream())
                    {
                        fs.CopyTo(msTemp);
                        templateBytes = msTemp.ToArray();
                    }
                }

                using var ms = new MemoryStream(templateBytes);
                using var doc = new XWPFDocument(ms);


                // 1. 处理循环填充 (支持 List<DynamicCaseData> 上下文)
                ProcessLoops(doc, dataContext);

                // 2. 处理单值填充
                // 如果是 List 上下文，取第一条数据作为单值填充源
                DynamicCaseData singleData = null;
                if (dataContext is DynamicCaseData d) singleData = d;
                else if (dataContext is List<DynamicCaseData> list && list.Count > 0) singleData = list[0];
                
                if (singleData != null)
                {
                    ProcessSingles(doc, singleData);
                }

                // 3. 应用只读保护 (如果需要)
                if (isReadOnly)
                {
                    // 设置只读保护（密码为空或随机，防止轻易编辑）
                    // EnforceReadonlyProtection API in NPOI 2.x for XWPF might slightly differ based on version.
                    // Assuming .NET Core NPOI port.
                    
                    // 方法1: 使用 EnforceReadonlyProtection (if available)
                    // doc.EnforceReadonlyProtection("PreviewOnly", NPOI.POIFS.Crypt.HashAlgorithm.sha1);
                    
                    // 方法2: 简单设置 Setting
                    /*
                    if (doc.GetCTDocument().settings == null) doc.GetCTDocument().AddNewSettings();
                    var settings = doc.GetCTDocument().settings;
                    // ... complex XML manipulation ...
                    */
                    
                    // 鉴于NPOI版本差异，这里使用 EnforceReadonlyProtection 如果存在，或 EnforceUpdateFields 替代干扰?
                    // 使用最通用的 doc.EnforceReadonlyProtection()
                     try 
                     {
                         doc.EnforceReadonlyProtection("PreviewLock", NPOI.POIFS.Crypt.HashAlgorithm.sha1);
                     }
                     catch
                     {
                         // Fallback or ignore if not supported
                     }
                }

                using (var outFile = File.Create(outputFilePath))
                {
                    doc.Write(outFile);
                }
                
                // 4. 文件系统级只读 (双重保障)
                if (isReadOnly && File.Exists(outputFilePath))
                {
                    try 
                    {
                        File.SetAttributes(outputFilePath, File.GetAttributes(outputFilePath) | FileAttributes.ReadOnly);
                    }
                    catch { }
                }

                return true;
            }
            catch (Exception ex)
            {
                LogHelper.Error("生成文书失败", ex);
                return false;
            }
        }

        /// <summary>
        /// 批量生成文书
        /// </summary>
        public static int BatchGenerateDocumentsFromExcel(string templatePath, string outputDir, List<Dictionary<string, string>> excelRowDatas, string caseKeyField = "案件编号")
        {
            if (string.IsNullOrEmpty(templatePath) || !File.Exists(templatePath)) 
                return 0;

            if (!Directory.Exists(outputDir)) Directory.CreateDirectory(outputDir);

            byte[] templateBytes;
            try 
            { 
                 using (var fs = new FileStream(templatePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                 {
                     using (var msTemp = new MemoryStream())
                     {
                         fs.CopyTo(msTemp);
                         templateBytes = msTemp.ToArray();
                     }
                 }
            }
            catch (Exception ex) { LogHelper.Error($"读取模板失败：{ex.Message}"); return 0; }

            int successCount = 0;


            foreach (var rowData in excelRowDatas)
            {
                try
                {
                    var caseData = new DynamicCaseData();
                    foreach (var kv in rowData)
                    {
                        caseData.CaseInfo[kv.Key] = kv.Value;
                    }
                    string keyVal = rowData.ContainsKey(caseKeyField) ? rowData[caseKeyField] : Guid.NewGuid().ToString("N");
                    caseData.CaseId = keyVal;

                    using var ms = new MemoryStream(templateBytes);
                    using var doc = new XWPFDocument(ms);

                    ProcessLoops(doc, caseData);
                    ProcessSingles(doc, caseData);

                    string fileName = $"{keyVal}_{DateTime.Now:yyyyMMddHHmmss}.docx";
                    string fullPath = Path.Combine(outputDir, fileName);

                    using (var outFile = File.Create(fullPath))
                    {
                        doc.Write(outFile);
                    }
                    successCount++;
                }
                catch (Exception ex)
                {
                    LogHelper.Error($"生成单条文书失败", ex);
                }
            }

            return successCount;
        }

        #endregion

        #region 逻辑实现 - 循环处理

        private static void ProcessLoops(XWPFDocument doc, object dataContext)
        {
            // 处理表格循环
            foreach (var table in doc.Tables)
            {
                ProcessTableLoops(table, dataContext);
            }
            
            // 处理正文段落循环
            ProcessBodyLoops(doc, dataContext);
        }

        private static void ProcessBodyLoops(XWPFDocument doc, object dataContext)
        {
            // 扫描文档Body元素
            // 必须反复扫描，因为插入操作会改变集合
            bool modificationHappened = true;
            while (modificationHappened)
            {
                modificationHappened = false;
                var bodyElements = doc.BodyElements; // 重新获取
                
                for (int i = 0; i < bodyElements.Count; i++)
                {
                    if (bodyElements[i].ElementType != BodyElementType.PARAGRAPH) continue;

                    var para = (XWPFParagraph)bodyElements[i];
                    var startMatch = Regex.Match(para.ParagraphText, @"\{\{" + LoopStartPattern + @"\}\}");
                    if (startMatch.Success)
                    {
                        string key = startMatch.Groups[1].Value;
                        int startIndex = i;
                        int endIndex = -1;

                        // 寻找结束标记
                        for (int j = i + 1; j < bodyElements.Count; j++)
                        {
                            var e = bodyElements[j];
                            if (e.ElementType == BodyElementType.PARAGRAPH)
                            {
                                if (((XWPFParagraph)e).ParagraphText.Contains($"{{{{LOOP_{key}_END}}}}"))
                                {
                                    endIndex = j;
                                    break;
                                }
                            }
                        }

                        if (endIndex != -1)
                        {
                            var listData = GetListDataFromContext(dataContext, key);

                            // 1. 确定模板范围 (不包含Start和End标记行)
                            // 修正：实际上通常包含在Start/End行中间的内容。
                            // 这里定义：Start标记所在行 到 End标记所在行 之间的内容为模板。
                            int templateStart = startIndex + 1;
                            int templateEnd = endIndex - 1;
                            
                            // 暂存模板元素引用
                            var templateElems = new List<IBodyElement>();
                            for(int k=templateStart; k<=templateEnd; k++)
                            {
                                templateElems.Add(bodyElements[k]);
                            }

                            // 2. 执行插入
                            // 插入位置：Start标记所在位置 (即替换整个块)
                            int insertPos = startIndex;
                            
                            if (listData != null)
                            {
                                foreach (var dataItem in listData)
                                {
                                    foreach (var tmplElem in templateElems)
                                    {
                                        if (tmplElem.ElementType == BodyElementType.PARAGRAPH)
                                        {
                                            // 创建新段落 (默认在文档末尾)
                                            var newPara = doc.CreateParagraph();
                                            CopyParagraph((XWPFParagraph)tmplElem, newPara);
                                            ReplaceValuesInParagraph(newPara, dataItem);
                                            
                                            // 移动到指定位置
                                            MoveBodyElement(doc, newPara, insertPos);
                                        }
                                        else if (tmplElem.ElementType == BodyElementType.TABLE)
                                        {
                                            var newTable = doc.CreateTable();
                                            CopyTable((XWPFTable)tmplElem, newTable);
                                            foreach (var row in newTable.Rows) ReplaceValuesInRow(row, dataItem);
                                            
                                            MoveBodyElement(doc, newTable, insertPos);
                                        }
                                        insertPos++;
                                    }
                                }
                            }

                            // 3. 删除原始 Loop 块 (包含 Start, Template, End)
                            // 此时原始块已被推到 insertPos 之后
                            // CRITIAL: 清空待删除段落的文本，以防止在 BodyElements 缓存未更新的情况下再次匹配到 Loop 标记导致死循环
                            // 旧的 Wrapper 仍然保留在 BodyElements 中 (因为我们只从 XML 中删除了)
                            
                            // 注意：insertPos 指向的是原始块开始的位置（因为它被上面的插入推到了后面）
                            // 但是 BodyElements 列表的顺序并没有因为 MoveBodyElement 改变 (如果我们无法刷新它)
                            // BodyElements 里的 Wrapper 顺序仍然是 [Start, Tmpl, End, New1, New2...] (如果 New 被添加到了末尾)
                            
                            // 修正逻辑：
                            // 如果 `doc.BodyElements` 是 STALE 的，它还是 [Start(old), Tmpl(old), End(old), New1(latest)...]
                            // 我们的循环遍历的是 STALE 列表。此时 startIndex 指向 Start。
                            // endIndex 指向 End。
                            // 我们生成的 NEW 元素虽然在 XML 中被移到了前面，但在 BODYELEMENTS 列表中是在最后。
                            
                            // 因此，我们必须清除 startIndex 到 endIndex 范围内的 Wrapper 内容。
                            for(int k = startIndex; k <= endIndex; k++)
                            {
                                if(bodyElements[k] is XWPFParagraph p)
                                {
                                    // 清空文本 runs
                                    while(p.Runs.Count > 0) p.RemoveRun(0);
                                    // 如果有 Text (rare), reset
                                    // p.GetCTP().SetValue(""); // Not easy on high level
                                }
                                // Table wrapper? Table doesn't trigger regex match in this loop.
                            }

                            // 删除位置: insertPos
                            // 注意：removeBodyElements 是基于 XML 索引操作的。
                            // 当我们插入了 N 个新元素并移到 startIndex 处，原来的 XML 元素确实被推到了 startIndex + N (即 current insertPos)。
                            // 所以从 insertPos 开始移除 countToRemove 个元素，确实移除了旧的 XML 节点。
                            // 这是正确的。
                            
                            int countToRemove = endIndex - startIndex + 1;
                            RemoveBodyElements(doc, insertPos, countToRemove);

                            modificationHappened = true;
                            break; // 重新扫描

                        }
                    }
                }
            }
        }

        private static void ProcessTableLoops(XWPFTable table, object dataContext)
        {
            for (int i = 0; i < table.Rows.Count; i++)
            {
                var row = table.Rows[i];
                string rowText = GetRowText(row);

                var startMatch = Regex.Match(rowText, @"\{\{" + LoopStartPattern + @"\}\}");
                if (startMatch.Success)
                {
                    string key = startMatch.Groups[1].Value;
                    int startIndex = i;
                    int endIndex = -1;

                    // 寻找结束行
                    for (int j = i; j < table.Rows.Count; j++)
                    {
                        var loopEndMatch = Regex.Match(GetRowText(table.Rows[j]), @"\{\{LOOP_" + key + @"_END\}\}");
                        if (loopEndMatch.Success)
                        {
                            endIndex = j;
                            break;
                        }
                    }

                    if (endIndex != -1)
                    {
                        var listData = GetListDataFromContext(dataContext, key);
                        int insertPos = endIndex + 1;
                        int templateRowCount = endIndex - startIndex + 1;
                        
                        // 捕获模板行（引用）以便复制
                        // 注意：我们不能一边插入一边读取模板，如果在同一表中，插入会改变索引
                        // 但我们记录了 startIndex 和 endIndex，这范围内的就是模板。
                        // 我们需要 "Manual Clone" 这一组行
                        
                        // 先移除标记文本（在生成前，为了避免复制后的行也带标记）
                        // 不，应该生成后再移除？或者生成时替换。
                        // 策略：保留模板不动，最后删除。
                        
                        if (listData != null && listData.Count > 0)
                        {
                            foreach (var item in listData)
                            {
                                for (int r = startIndex; r <= endIndex; r++)
                                {
                                    var sourceRow = table.GetRow(r);
                                    var newRow = table.InsertNewTableRow(insertPos);
                                    CopyRow(sourceRow, newRow);
                                    
                                    // 替换变量
                                    ReplaceValuesInRow(newRow, item);
                                    
                                    // 移除 Loop 标记 (New Row)
                                    ReplaceTextInRow(newRow, startMatch.Value, "");
                                    ReplaceTextInRow(newRow, $"{{{{LOOP_{key}_END}}}}", "");
                                    
                                    insertPos++;
                                }
                            }
                        }

                        // 删除原始模板行
                        for (int k = 0; k < templateRowCount; k++)
                        {
                            table.RemoveRow(startIndex);
                        }
                        
                        // 修正循环索引
                        int addedRows = (listData?.Count ?? 0) * templateRowCount;
                        i = startIndex + addedRows - 1;
                    }
                }
            }
        }

        #endregion

        #region 逻辑实现 - 单值处理

        private static void ProcessSingles(XWPFDocument doc, DynamicCaseData caseData)
        {
            var dict = caseData.CaseInfo;

            foreach (var para in doc.Paragraphs)
            {
                ReplaceValuesInParagraph(para, dict);
            }

            foreach (var table in doc.Tables)
            {
                foreach (var row in table.Rows)
                {
                   ReplaceValuesInRow(row, dict);
                }
            }
        }

        #endregion

        #region 辅助方法

        private static string GetRowText(XWPFTableRow row)
        {
             StringBuilder sb = new StringBuilder();
             foreach(var cell in row.GetTableCells())
             {
                 sb.Append(cell.GetText());
             }
             return sb.ToString();
        }

        private static void ReplaceTextInRow(XWPFTableRow row, string placeholder, string newVal)
        {
            foreach(var cell in row.GetTableCells())
            {
                foreach(var para in cell.Paragraphs)
                {
                    ReplaceTextInParagraph(para, placeholder, newVal);
                }
            }
        }
        
        private static List<Dictionary<string, object>>? GetListData(DynamicCaseData caseData, string key)
        {
            if (!caseData.CaseInfo.ContainsKey(key)) return null;
            var val = caseData.CaseInfo[key];
            if (val is List<Dictionary<string, object>> list) return list;
            if (val is string str && str.Trim().StartsWith("[") && str.Trim().EndsWith("]"))
            {
                try { return JsonConvert.DeserializeObject<List<Dictionary<string, object>>>(str); } catch { }
            }
            return null;
        }

        // 新增辅助方法：统一从上下文获取 List 数据
        private static List<IDictionary<string, object>>? GetListDataFromContext(object context, string loopKey)
        {
            // 圆见 1: 上下文本身就是列表 (对应 "填充人数" 分割后的 Chunk)
            // 此时忽略 loopKey (或者 loopKey 仅作为标识，列表直接使用)
            // 需求："识别 Loop 标记...按分割后的子数组条数循环填充"
            // 这意味着 loopKey 即使是 Person，如果传入的是 List<Case>，就用 List<Case> 填充
            // 但如果文档有多个不同 Loop 呢？"同一模板中单个循环标识仅对应一组 Excel 数据列"
            // 这暗示通常只有一个主循环，或者多个循环都指代同一组数据。
            if (context is List<DynamicCaseData> caseList)
            {
                // 将 DynamicCaseData 列表转换为 IDictionary 列表供替换使用
                return caseList.Select(c => c.CaseInfo).ToList<IDictionary<string, object>>();
            }

            // 圆见 2: 上下文是单个 DynamicCaseData (旧逻辑，对象内包含 List 属性)
            if (context is DynamicCaseData singleCase)
            {
                // 尝试从 CaseInfo 中获取名为 loopKey 的列表
                if (singleCase.CaseInfo.TryGetValue(loopKey, out var val))
                {
                     if (val is List<Dictionary<string, object>> list) 
                        return list.Cast<IDictionary<string, object>>().ToList(); // 转换类型适配
                     
                     // 支持 JSON 字符串
                     if (val is string str && str.Trim().StartsWith("[") && str.Trim().EndsWith("]"))
                     {
                         try { return JsonConvert.DeserializeObject<List<Dictionary<string, object>>>(str)?.Cast<IDictionary<string, object>>().ToList(); } catch { }
                     }

                }
            }
            
            LogHelper.Warn($"未找到循环数据：Key={loopKey}, ContextType={context?.GetType().Name}");
            return null;
        }

        private static void ReplaceValuesInRow(XWPFTableRow row, IDictionary<string, object> data)
        {
            foreach (var cell in row.GetTableCells())
            {
                foreach (var para in cell.Paragraphs)
                {
                    ReplaceValuesInParagraph(para, data);
                }
            }
        }

        private static void ReplaceValuesInParagraph(XWPFParagraph para, IDictionary<string, object> data)
        {
            if (data == null) return;
            string text = para.ParagraphText;
            // 宽松检查，因为可能包含其他标记
            if (string.IsNullOrEmpty(text) || !text.Contains("{{")) return;

            // 优先替换最长匹配的键，避免部分匹配问题
            // 但这里遍历顺序不可控，且NPOI替换会改变ParagraphText 
            // 建议：收集所有匹配项，再一次性替换? NPOI Run操作比较复杂。
            // 维持现状，但添加Log调试

            foreach (var kv in data)
            {
                string key = kv.Key;
                string placeholder = $"{{{{{key}}}}}";
                
                // 检查段落是否包含该占位符
                if (text.Contains(placeholder))
                {
                    string val = kv.Value?.ToString() ?? "";
                    // 调试日志：发现占位符
                    // LogHelper.Debug($"Replacing [{placeholder}] with [{val}] in paragraph: {text.Substring(0, Math.Min(text.Length, 20))}...");
                    
                    ReplaceTextInParagraph(para, placeholder, val);
                    
                    // 更新text以便后续匹配（虽然ReplaceTextInParagraph内部已经修改了runs，ParagraphText属性应该是计算出来的）
                    text = para.ParagraphText;
                }
            }
        }

        private static void ReplaceTextInParagraph(XWPFParagraph para, string placeholder, string newText)
        {
            var text = para.ParagraphText;
            if (!text.Contains(placeholder)) return;

            // 1. 尝试在单个Run中查找 (Simple replace)
            foreach (var run in para.Runs)
            {
                string runText = run.Text;
                if (!string.IsNullOrEmpty(runText) && runText.Contains(placeholder))
                {
                    run.SetText(runText.Replace(placeholder, newText), 0);
                    return;
                }
            }

            // 2. 跨 Run 查找 (Complex replace)
            // 这种情况下，占位符被分割到了多个Run中，比如 "{{" 在 Run1, "Key" 在 Run2, "}}" 在 Run3
            // 简单处理：重置第一个Run为全文替换后的结果，清空其他Runs。
            // 缺点：会丢失段落中其他的格式（如颜色加粗不一致的部分）。
            // 但为了保证替换成功，这是常用妥协方案。

            string fullText = para.ParagraphText;
            string replacedText = fullText.Replace(placeholder, newText);
            
            // 如果替换无效（理论上Contains已检查），退出
            if (fullText == replacedText) return;

            if (para.Runs.Count > 0)
            {
                var firstRun = para.Runs[0];
                firstRun.SetText(replacedText, 0);
                
                // 必须倒序删除
                for (int i = para.Runs.Count - 1; i > 0; i--)
                {
                    para.RemoveRun(i);
                }
            }
            else
            {
                // 无Run但有Text? 可能是NPOI Text缓存。创建新Run。
                para.CreateRun().SetText(replacedText);
            }
        }

        #endregion

        #region 样式复制辅助方法 (Manual Deep Copy)

        private static void CopyRow(XWPFTableRow source, XWPFTableRow target)
        {
            // 复制行属性
            if (source.GetCTRow().trPr != null)
            {
                target.GetCTRow().trPr = source.GetCTRow().trPr;
            }

            // 复制单元格
            // InsertNewTableRow 可能会创建一个空单元格或无单元格，视版本而定。
            // 检查 Target 单元格数
            var sourceCells = source.GetTableCells();
            var targetCells = target.GetTableCells();

            // 如果 Target 已经有默认单元格（通常有1个），先利用它，不足的创建
            // 如果 Target 为空，全部创建
            // 为简单起见，清除 Target 现有（如果有 API），或者覆盖
            
            // 多数情况 InsertNewTableRow(i) 创建的行包含 1 个空单元格
            // 我们需要对齐数量
            
            for (int i = 0; i < sourceCells.Count; i++)
            {
                XWPFTableCell? targetCell = null;
                if (i < targetCells.Count) targetCell = targetCells[i];
                else targetCell = target.AddNewTableCell();

                CopyCell(sourceCells[i], targetCell);
            }
        }

        private static void CopyCell(XWPFTableCell source, XWPFTableCell target)
        {
            // 复制属性
            if (source.GetCTTc().tcPr != null)
            {
                target.GetCTTc().tcPr = source.GetCTTc().tcPr;
            }

            // 清除 Target 默认段落（通常新建 Cell 会带一个空段落）
            // NPOI 中 RemoveParagraph 需要 index
            while(target.Paragraphs.Count > 0)
            {
                target.RemoveParagraph(0);
            }

            // 复制段落
            foreach (var para in source.Paragraphs)
            {
                var newPara = target.AddParagraph();
                CopyParagraph(para, newPara);
            }
        }

        private static void CopyParagraph(XWPFParagraph source, XWPFParagraph target)
        {
            // 复制段落属性
            if (source.GetCTP().pPr != null)
            {
                target.GetCTP().pPr = source.GetCTP().pPr;
            }

            // 复制 Runs
            foreach (var run in source.Runs)
            {
                var newRun = target.CreateRun();
                CopyRun(run, newRun);
            }
        }

        private static void CopyRun(XWPFRun source, XWPFRun target)
        {
            // 复制 Run 属性
            if (source.GetCTR().rPr != null)
            {
                target.GetCTR().rPr = source.GetCTR().rPr;
            }
            target.SetText(source.Text, 0);
        }

        private static void CopyTable(XWPFTable source, XWPFTable target)
        {
             // 复制表格属性
             if (source.GetCTTbl().tblPr != null)
             {
                 target.GetCTTbl().tblPr = source.GetCTTbl().tblPr;
             }
             
             // 复制行
             // target 默认有1行
             target.RemoveRow(0);
             
             foreach(var row in source.Rows)
             {
                 var newRow = target.CreateRow();
                 CopyRow(row, newRow);
             }
        }

        private static void MoveBodyElement(XWPFDocument doc, IBodyElement elem, int index)
        {
             var body = doc.Document.body;
             // NPOI 2.x CT_Body Items is a List<object> (ArrayList in older versions) of CT_P and CT_Tbl
             // Use dynamic to access Items property on CT_Body
             dynamic dBody = body;
             System.Collections.IList items = dBody.Items;
             
             object? ctObj = null;
             if(elem is XWPFParagraph p) ctObj = p.GetCTP();
             else if(elem is XWPFTable t) ctObj = t.GetCTTbl();
             
             if (ctObj != null && items.Contains(ctObj))
             {
                 items.Remove(ctObj); // Remove from current location (usually end)
                 if (index >= 0 && index <= items.Count)
                 {
                     items.Insert(index, ctObj);
                 }
                 else
                 {
                     items.Add(ctObj); // Fallback
                 }
             }
        }

        private static void RemoveBodyElements(XWPFDocument doc, int startIndex, int count)
        {
             var body = doc.Document.body;
             dynamic dBody = body;
             System.Collections.IList items = dBody.Items;
             
             // Remove count items starting from startIndex
             // Since removing shifts indices, we remove at startIndex repeatedly
             for(int i=0; i<count; i++)
             {
                 if(startIndex < items.Count)
                 {
                     items.RemoveAt(startIndex);
                 }
             }
        }
        
        // ICursor 接口声明，避免引用错误 (NPOI.XWPF.UserModel 没有直接暴露 ICursor 接口，而是 XmlCursor?)

        // 在 C# NPOI 中，newCursor 返回的是 XmlCursor 类型 (System.Xml.XmlObject 的一部分?)
        // 实际上 NPOI 2.x .NET 是基于 OpenXmlFormats，底层 xmlbean 只有部分暴露
        // 如果无法使用 Cursor，ProcessBodyLoops 只能暂时搁置逻辑
        private interface ICursor 
        { 
             void toNextToken(); 
             // ... dummy interface for code structure if needed, but logic commented out
        } 
        // 实际逻辑中已使用 dynamic 或 object 替代，如果无法编译则删除相关行

        #endregion

        public static readonly string SystemTemplateDir = LocalStorageHelper._systemTemplateDir;
        public static readonly string CustomTemplateDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "LegalDocumentTemplates");

        static WordService()
        {
            if (!Directory.Exists(CustomTemplateDir)) Directory.CreateDirectory(CustomTemplateDir);
        }
    }
}
