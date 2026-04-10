using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace WpfApp1.Services
{
    /// <summary>解析 CREATE TABLE DDL 语句，提取字段元数据。</summary>
    public static class DdlParser
    {
        /// <summary>根据 DDL 特征关键词推断数据库类型，无法判断时返回 null。</summary>
        public static DvDbType? DetectDbType(string ddl)
        {
            if (string.IsNullOrWhiteSpace(ddl)) return null;

            int pg = 0, ss = 0;

            // ── PostgreSQL 强信号 ──
            if (Regex.IsMatch(ddl, @"\buuid\b", RegexOptions.IgnoreCase)) pg += 3;
            if (Regex.IsMatch(ddl, @"\b(big)?serial\b", RegexOptions.IgnoreCase)) pg += 3;
            if (Regex.IsMatch(ddl, @"\bboolean\b", RegexOptions.IgnoreCase)) pg += 2;
            if (Regex.IsMatch(ddl, @"\bjsonb\b", RegexOptions.IgnoreCase)) pg += 3;
            if (Regex.IsMatch(ddl, @"\btimestamptz\b", RegexOptions.IgnoreCase)) pg += 3;
            if (Regex.IsMatch(ddl, @"\bcharacter\s+varying\b", RegexOptions.IgnoreCase)) pg += 2;
            if (Regex.IsMatch(ddl, @"\bcitext\b", RegexOptions.IgnoreCase)) pg += 3;
            if (ddl.Contains("::")) pg += 2;

            // ── SQL Server 强信号 ──
            if (Regex.IsMatch(ddl, @"\buniqueidentifier\b", RegexOptions.IgnoreCase)) ss += 3;
            if (Regex.IsMatch(ddl, @"\bn(var)?char\b", RegexOptions.IgnoreCase)) ss += 2;
            if (Regex.IsMatch(ddl, @"\bntext\b", RegexOptions.IgnoreCase)) ss += 3;
            if (Regex.IsMatch(ddl, @"\bdatetime2\b", RegexOptions.IgnoreCase)) ss += 3;
            if (Regex.IsMatch(ddl, @"\bsmalldatetime\b", RegexOptions.IgnoreCase)) ss += 3;
            if (Regex.IsMatch(ddl, @"\bdatetimeoffset\b", RegexOptions.IgnoreCase)) ss += 3;
            if (Regex.IsMatch(ddl, @"\b(small)?money\b", RegexOptions.IgnoreCase)) ss += 3;
            if (Regex.IsMatch(ddl, @"\btinyint\b", RegexOptions.IgnoreCase)) ss += 2;
            if (Regex.IsMatch(ddl, @"\bIDENTITY\s*\(", RegexOptions.IgnoreCase)) ss += 2;
            if (ddl.Contains('[') && ddl.Contains(']')) ss += 2;

            if (pg == 0 && ss == 0) return null;
            if (pg > ss) return DvDbType.PostgreSql;
            if (ss > pg) return DvDbType.SqlServer;
            return null; // 无法确定
        }

        /// <summary>从 DDL 中提取表名（去掉 schema 前缀和引号）。</summary>
        public static string? ExtractTableName(string ddl)
        {
            if (string.IsNullOrWhiteSpace(ddl)) return null;
            var m = Regex.Match(ddl,
                @"\bCREATE\s+TABLE\s+(?:IF\s+NOT\s+EXISTS\s+)?(?:[\w`""\[\]]+\.)?[`""\[]?(\w+)[`""\]]?\s*\(",
                RegexOptions.IgnoreCase);
            return m.Success ? m.Groups[1].Value : null;
        }

        public static List<DvTargetColumn> Parse(string ddl, DvDbType dbType)
        {
            if (string.IsNullOrWhiteSpace(ddl))
                throw new InvalidOperationException("DDL 语句为空。");

            // 去掉行注释
            ddl = Regex.Replace(ddl, @"--[^\r\n]*", " ");
            // 去掉块注释
            ddl = Regex.Replace(ddl, @"/\*.*?\*/", " ", RegexOptions.Singleline);

            var block = ExtractColumnBlock(ddl);
            if (block == null)
                throw new InvalidOperationException("未找到字段定义块（找不到括号内内容）。");

            var result = new List<DvTargetColumn>();
            int ordinal = 1;

            foreach (var part in SplitTopLevel(block))
            {
                var trimmed = part.Trim();
                if (string.IsNullOrWhiteSpace(trimmed)) continue;

                // 跳过表级约束
                var upper = trimmed.ToUpperInvariant();
                if (upper.StartsWith("PRIMARY") || upper.StartsWith("UNIQUE") ||
                    upper.StartsWith("INDEX") || upper.StartsWith("CONSTRAINT") ||
                    upper.StartsWith("KEY") || upper.StartsWith("CHECK"))
                    continue;

                var col = ParseColumn(trimmed, ordinal, dbType);
                if (col != null)
                {
                    result.Add(col);
                    ordinal++;
                }
            }

            if (result.Count == 0)
                throw new InvalidOperationException("解析后字段数为 0，请检查 DDL 格式。");

            return result;
        }

        // ── 提取最外层括号内的内容 ────────────────────────────
        private static string? ExtractColumnBlock(string ddl)
        {
            int start = ddl.IndexOf('(');
            if (start < 0) return null;

            int depth = 0;
            int end = -1;
            for (int i = start; i < ddl.Length; i++)
            {
                if (ddl[i] == '(') depth++;
                else if (ddl[i] == ')')
                {
                    depth--;
                    if (depth == 0) { end = i; break; }
                }
            }
            if (end < 0) return null;
            return ddl.Substring(start + 1, end - start - 1);
        }

        // ── 按顶层逗号分割（不处理括号内的逗号）────────────────
        private static IEnumerable<string> SplitTopLevel(string block)
        {
            int depth = 0;
            int start = 0;
            for (int i = 0; i < block.Length; i++)
            {
                char c = block[i];
                if (c == '(') depth++;
                else if (c == ')') depth--;
                else if (c == ',' && depth == 0)
                {
                    yield return block.Substring(start, i - start);
                    start = i + 1;
                }
            }
            if (start < block.Length)
                yield return block.Substring(start);
        }

        // ── 解析单个字段定义 ─────────────────────────────────
        private static DvTargetColumn? ParseColumn(string def, int ordinal, DvDbType dbType)
        {
            // 去掉 DEFAULT xxx, IDENTITY, COLLATE xxx（用正则把这些子句剥离）
            def = Regex.Replace(def, @"\bDEFAULT\b\s+\S+", "", RegexOptions.IgnoreCase);
            def = Regex.Replace(def, @"\bIDENTITY\s*(\(\s*\d+\s*,\s*\d+\s*\))?", "", RegexOptions.IgnoreCase);
            def = Regex.Replace(def, @"\bCOLLATE\b\s+\S+", "", RegexOptions.IgnoreCase);
            def = Regex.Replace(def, @"\bGENERATED\b.*", "", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            def = def.Trim();

            // 字段名（可能带 ` 或 " 或 [ ] 引号）
            var nameMatch = Regex.Match(def, @"^[`""\[]?(\w+)[`""\]]?\s+");
            if (!nameMatch.Success) return null;
            string colName = nameMatch.Groups[1].Value;
            string rest = def.Substring(nameMatch.Length).Trim();

            // 类型（含可选括号参数）
            var typeMatch = Regex.Match(rest, @"^([\w\s]+?)(\s*\(\s*([^)]*)\s*\))?(?=\s|$)", RegexOptions.IgnoreCase);
            if (!typeMatch.Success) return null;

            string rawType = typeMatch.Groups[1].Value.Trim().ToLower();
            string? paramStr = typeMatch.Groups[3].Success ? typeMatch.Groups[3].Value.Trim() : null;

            // 解析括号参数
            int? maxLen = null;
            int? precision = null;
            int? scale = null;

            if (paramStr != null)
            {
                if (paramStr.Equals("max", StringComparison.OrdinalIgnoreCase))
                {
                    maxLen = -1;
                }
                else
                {
                    var parts = paramStr.Split(',');
                    if (parts.Length == 1 && int.TryParse(parts[0].Trim(), out int p1))
                    {
                        // 字符串类型用 maxLen，数值类型用 precision
                        if (IsStringType(rawType))
                            maxLen = p1;
                        else
                            precision = p1;
                    }
                    else if (parts.Length == 2 &&
                             int.TryParse(parts[0].Trim(), out int pre) &&
                             int.TryParse(parts[1].Trim(), out int sc))
                    {
                        precision = pre;
                        scale = sc;
                    }
                }
            }

            // CHARACTER VARYING 等多词类型修正
            rawType = NormalizeRawType(rawType);

            // 是否可空
            string afterType = rest.Substring(typeMatch.Length).ToUpper();
            bool isNullable = !afterType.Contains("NOT NULL");

            var normalized = SchemaNormalizer.Normalize(rawType, dbType);

            return new DvTargetColumn
            {
                OrdinalPosition = ordinal,
                ColumnName = colName,
                OriginalDataType = rawType,
                NormalizedType = normalized,
                MaxLength = maxLen,
                NumericPrecision = precision,
                NumericScale = scale,
                IsNullable = isNullable,
                DatabaseType = dbType
            };
        }

        private static string NormalizeRawType(string raw)
        {
            // 多词类型归一
            raw = Regex.Replace(raw, @"\s+", " ");
            return raw switch
            {
                "character varying" => "varchar",
                "double precision" => "float",
                "timestamp without time zone" => "timestamp",
                "timestamp with time zone" => "timestamptz",
                "time without time zone" => "time",
                "time with time zone" => "timetz",
                _ => raw
            };
        }

        private static bool IsStringType(string rawType) =>
            rawType is "char" or "varchar" or "nchar" or "nvarchar" or
                       "character" or "character varying" or "binary" or "varbinary";
    }
}
