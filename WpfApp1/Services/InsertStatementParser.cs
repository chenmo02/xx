using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace WpfApp1.Services
{
    /// <summary>解析 INSERT INTO 语句，提取字段列表和数据行。</summary>
    public static class InsertStatementParser
    {
        public static (List<string> Headers, List<IReadOnlyList<string?>> Rows, string? Warning) Parse(string sql)
        {
            if (string.IsNullOrWhiteSpace(sql))
                throw new InvalidOperationException("SQL 语句为空。");

            // 去掉注释
            sql = Regex.Replace(sql, @"--[^\r\n]*", " ");
            sql = Regex.Replace(sql, @"/\*.*?\*/", " ", RegexOptions.Singleline);

            // 提取第一条 INSERT 的字段列表
            var headerMatch = Regex.Match(sql,
                @"\bINSERT\s+INTO\s+[\w\.\[\]`""]+\s*\(([^)]+)\)\s*VALUES?\b",
                RegexOptions.IgnoreCase | RegexOptions.Singleline);

            if (!headerMatch.Success)
                throw new InvalidOperationException("未找到有效的 INSERT INTO 语句（需要显式字段列表）。");

            var headers = ParseIdentifierList(headerMatch.Groups[1].Value);
            if (headers.Count == 0)
                throw new InvalidOperationException("字段列表解析结果为空。");

            // 找所有 VALUES 后的元组
            var rows = new List<IReadOnlyList<string?>>();
            string? warning = null;
            bool firstMatch = true;

            // 逐条处理每个 INSERT INTO ... VALUES (...)
            var allInserts = Regex.Matches(sql,
                @"\bINSERT\s+INTO\s+[\w\.\[\]`""]+\s*\(([^)]+)\)\s*VALUES?\s*([\s\S]+?)(?=;|\bINSERT\b|$)",
                RegexOptions.IgnoreCase);

            foreach (Match m in allInserts)
            {
                var colsPart = m.Groups[1].Value;
                var valuesPart = m.Groups[2].Value.Trim().TrimEnd(';').Trim();

                if (!firstMatch)
                {
                    var thisHeaders = ParseIdentifierList(colsPart);
                    if (!HeadersMatch(headers, thisHeaders) && warning == null)
                        warning = "检测到多条 INSERT 字段列表不一致，将以第一条为准。";
                    firstMatch = false;
                }
                else
                {
                    firstMatch = false;
                }

                // valuesPart 可能是多个 (v1,v2,...), (v3,v4,...) 用逗号或换行分隔
                foreach (var tuple in ExtractTuples(valuesPart))
                {
                    var vals = ParseTuple(tuple, headers.Count);
                    rows.Add(vals);
                }
            }

            if (rows.Count == 0)
                throw new InvalidOperationException("未解析到任何数据行，请检查 VALUES 部分格式。");

            return (headers, rows, warning);
        }

        // ── 解析以逗号分隔的标识符列表 ─────────────────────────
        private static List<string> ParseIdentifierList(string s)
        {
            var result = new List<string>();
            foreach (var part in s.Split(','))
            {
                var name = part.Trim().Trim('`', '"', '[', ']');
                if (!string.IsNullOrEmpty(name))
                    result.Add(name);
            }
            return result;
        }

        // ── 从 VALUES 段提取所有括号内的元组 ─────────────────────
        private static IEnumerable<string> ExtractTuples(string valuesPart)
        {
            int i = 0;
            while (i < valuesPart.Length)
            {
                // 找下一个 (
                while (i < valuesPart.Length && valuesPart[i] != '(') i++;
                if (i >= valuesPart.Length) break;

                int start = i + 1;
                int depth = 1;
                i++;
                bool inStr = false;
                while (i < valuesPart.Length && depth > 0)
                {
                    char c = valuesPart[i];
                    if (c == '\'' && !inStr) inStr = true;
                    else if (c == '\'' && inStr)
                    {
                        // 检查转义 ''
                        if (i + 1 < valuesPart.Length && valuesPart[i + 1] == '\'')
                            i++; // skip ''
                        else
                            inStr = false;
                    }
                    else if (!inStr)
                    {
                        if (c == '(') depth++;
                        else if (c == ')') depth--;
                    }
                    i++;
                }

                if (depth == 0)
                    yield return valuesPart.Substring(start, i - start - 1);
            }
        }

        // ── 解析单个元组内的值列表 ────────────────────────────
        private static List<string?> ParseTuple(string tuple, int expectedCount)
        {
            var result = new List<string?>();
            int i = 0;
            while (i <= tuple.Length)
            {
                // 跳过前导空白
                while (i < tuple.Length && char.IsWhiteSpace(tuple[i])) i++;
                if (i >= tuple.Length) break;

                if (tuple[i] == '\'')
                {
                    // 字符串值
                    i++;
                    var sb = new System.Text.StringBuilder();
                    while (i < tuple.Length)
                    {
                        if (tuple[i] == '\'')
                        {
                            if (i + 1 < tuple.Length && tuple[i + 1] == '\'')
                            {
                                sb.Append('\'');
                                i += 2;
                            }
                            else { i++; break; }
                        }
                        else { sb.Append(tuple[i]); i++; }
                    }
                    result.Add(sb.ToString());
                }
                else
                {
                    // 非字符串值（数字、NULL、布尔等）
                    int start = i;
                    while (i < tuple.Length && tuple[i] != ',') i++;
                    var raw = tuple.Substring(start, i - start).Trim();
                    result.Add(raw.Equals("NULL", StringComparison.OrdinalIgnoreCase) ? null : raw);
                }

                // 跳过逗号
                while (i < tuple.Length && char.IsWhiteSpace(tuple[i])) i++;
                if (i < tuple.Length && tuple[i] == ',') i++;
            }
            return result;
        }

        private static bool HeadersMatch(List<string> a, List<string> b)
        {
            if (a.Count != b.Count) return false;
            for (int i = 0; i < a.Count; i++)
                if (!string.Equals(a[i], b[i], StringComparison.OrdinalIgnoreCase))
                    return false;
            return true;
        }
    }
}
