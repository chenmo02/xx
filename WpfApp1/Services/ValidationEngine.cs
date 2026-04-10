using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace WpfApp1.Services
{
    public static class ValidationEngine
    {
        private const int MaxErrors = 10000;

        private static readonly string[] DateFormats =
            ["yyyy-MM-dd", "yyyy/MM/dd", "yyyyMMdd"];

        private static readonly string[] DateTimeFormats =
            ["yyyy-MM-dd HH:mm:ss", "yyyy/MM/dd HH:mm:ss",
             "yyyy-MM-dd HH:mm:ss.fff", "yyyy-MM-dd", "yyyy/MM/dd",
             "yyyyMMddHHmmss", "yyyyMMdd"];

        public static async Task<DvValidationResult> RunAsync(
            IReadOnlyList<DvTargetColumn> schema,
            DvSourceData data,
            IReadOnlyList<DvMappingRow> mappings,
            IReadOnlyList<string>? primaryKeyColumns = null,
            bool skipIntegerFormatErrors = false,
            IProgress<(int current, int total, int errors)>? progress = null,
            CancellationToken ct = default)
        {
            var start = DateTime.UtcNow;
            var issues = new List<DvIssue>();
            int processed = 0;
            int errorCount = 0;

            // 建索引：targetColumnName → headerIndex
            var colIndex = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < data.Headers.Count; i++)
                colIndex[data.Headers[i]] = i;

            // 主键列索引
            var pkIndices = new List<(string name, int index)>();
            if (primaryKeyColumns != null)
                foreach (var pk in primaryKeyColumns)
                    if (colIndex.TryGetValue(pk, out int idx))
                        pkIndices.Add((pk, idx));

            // 只处理 Source 映射，预解析列索引避免每行查字典
            var resolvedMappings = new List<(DvTargetColumn col, int sourceIndex, string sourceName)>();
            var schemaDict = schema.ToDictionary(c => c.ColumnName, StringComparer.OrdinalIgnoreCase);
            foreach (var m in mappings)
            {
                if (m.MappingType != DvMappingType.Source || m.SourceColumnName == null) continue;
                if (!schemaDict.TryGetValue(m.TargetColumnName, out var colMeta)) continue;
                if (!colIndex.TryGetValue(m.SourceColumnName, out int ci)) continue;
                resolvedMappings.Add((colMeta, ci, m.SourceColumnName));
            }

            // 预计算必填但未映射的字段列表（循环外只算一次）
            var unmappedRequired = new List<DvTargetColumn>();
            foreach (var col in schema.Where(c => !c.IsNullable))
            {
                var mapping = mappings.FirstOrDefault(m =>
                    string.Equals(m.TargetColumnName, col.ColumnName, StringComparison.OrdinalIgnoreCase));
                if (mapping == null || mapping.MappingType == DvMappingType.Ignore)
                {
                    if (mapping != null && mapping.IsAutoGenCandidate) continue;
                    unmappedRequired.Add(col);
                }
            }

            // 自适应上报间隔
            int reportInterval = Math.Max(100, data.Rows.Count / 200);

            await Task.Run(() =>
            {
                for (int rowIdx = 0; rowIdx < data.Rows.Count; rowIdx++)
                {
                    ct.ThrowIfCancellationRequested();

                    var row = data.Rows[rowIdx];
                    int rowNum = rowIdx + 2; // 行1是表头，数据从第2行起

                    // 构造主键标识
                    string? pkDisplay = null;
                    if (pkIndices.Count > 0)
                    {
                        pkDisplay = string.Join(", ", pkIndices.Select(pk =>
                            $"{pk.name}={((pk.index < row.Count ? row[pk.index] : null) ?? "NULL")}"));
                    }

                    foreach (var (colMeta, ci, srcName) in resolvedMappings)
                    {
                        string? value = ci < row.Count ? row[ci] : null;

                        foreach (var issue in ValidateCell(value, colMeta, rowNum, srcName, skipIntegerFormatErrors))
                        {
                            issue.PrimaryKeyDisplay = pkDisplay;
                            if (issue.Level == DvValidationLevel.Error) errorCount++;
                            issues.Add(issue);
                            if (issues.Count >= MaxErrors) goto Done;
                        }
                    }

                    // 必填未映射字段报错（已预计算）
                    foreach (var col in unmappedRequired)
                    {
                        errorCount++;
                        issues.Add(new DvIssue
                        {
                            RowNumber = rowNum,
                            PrimaryKeyDisplay = pkDisplay,
                            TargetColumnName = col.ColumnName,
                            TargetDataType = col.DisplayType,
                            Level = DvValidationLevel.Error,
                            ErrorType = "必填未映射",
                            ActualValue = null,
                            Message = $"必填字段 {col.ColumnName} 未映射或值为空"
                        });
                        if (issues.Count >= MaxErrors) goto Done;
                    }

                    processed++;
                    if (processed % reportInterval == 0)
                        progress?.Report((processed, data.Rows.Count, errorCount));
                }
                Done:;
            }, ct);

            // 最终上报一次
            progress?.Report((processed, data.Rows.Count, errorCount));

            return new DvValidationResult
            {
                Issues = issues,
                TotalRows = data.Rows.Count,
                ProcessedRows = processed,
                Elapsed = DateTime.UtcNow - start,
                WasCancelled = ct.IsCancellationRequested
            };
        }

        private static IEnumerable<DvIssue> ValidateCell(
            string? value, DvTargetColumn col, int rowNum, string? srcCol,
            bool skipIntegerFormatErrors = false)
        {
            // 1. 空值检查
            bool isEmpty = string.IsNullOrWhiteSpace(value);
            if (!col.IsNullable && isEmpty)
            {
                yield return Issue(rowNum, srcCol, col, DvValidationLevel.Error, "必填为空", value,
                    $"字段不可为空，实际值为空");
                yield break; // 空值无需继续类型校验
            }
            if (isEmpty) yield break;

            // 2. Trim 前后空格，不作为错误检测
            value = value!.Trim();

            switch (col.NormalizedType)
            {
                case DvNormalizedType.String:
                    foreach (var i in ValidateString(value!, col, rowNum, srcCol)) yield return i;
                    break;
                case DvNormalizedType.Integer:
                    if (!skipIntegerFormatErrors)
                        foreach (var i in ValidateInteger(value!, col, rowNum, srcCol)) yield return i;
                    break;
                case DvNormalizedType.Long:
                    if (!skipIntegerFormatErrors)
                        foreach (var i in ValidateLong(value!, col, rowNum, srcCol)) yield return i;
                    break;
                case DvNormalizedType.Decimal:
                    foreach (var i in ValidateDecimal(value!, col, rowNum, srcCol)) yield return i;
                    break;
                case DvNormalizedType.Date:
                    foreach (var i in ValidateDate(value!, col, rowNum, srcCol)) yield return i;
                    break;
                case DvNormalizedType.DateTime:
                    foreach (var i in ValidateDateTime(value!, col, rowNum, srcCol)) yield return i;
                    break;
                case DvNormalizedType.Boolean:
                    foreach (var i in ValidateBoolean(value!, col, rowNum, srcCol)) yield return i;
                    break;
                case DvNormalizedType.Guid:
                    foreach (var i in ValidateGuid(value!, col, rowNum, srcCol)) yield return i;
                    break;
                case DvNormalizedType.Json:
                    foreach (var i in ValidateJson(value!, col, rowNum, srcCol)) yield return i;
                    break;
                case DvNormalizedType.Unknown:
                    yield return Issue(rowNum, srcCol, col, DvValidationLevel.Warning, "未识别类型", value,
                        $"字段类型 {col.OriginalDataType} 未识别，仅做非空校验");
                    break;
            }
        }

        // ── 字符串 ────────────────────────────────────────────
        private static IEnumerable<DvIssue> ValidateString(string value, DvTargetColumn col, int row, string? src)
        {
            if (col.MaxLength.HasValue && col.MaxLength.Value > 0)
            {
                int len = value.Length;
                if (len > col.MaxLength.Value)
                    yield return Issue(row, src, col, DvValidationLevel.Error, "字符超长", value,
                        $"字符长度 {len} 超过限制 {col.MaxLength}");
                else if (col.DatabaseType == DvDbType.SqlServer &&
                         (col.OriginalDataType == "varchar" || col.OriginalDataType == "char") &&
                         value.Any(c => c > 127))
                    yield return Issue(row, src, col, DvValidationLevel.Warning, "字节可能超长", value,
                        "varchar 含多字节字符，实际字节占用可能超过限制");
            }
        }

        // ── 整数 ──────────────────────────────────────────────
        private static IEnumerable<DvIssue> ValidateInteger(string value, DvTargetColumn col, int row, string? src)
        {
            // 允许 "3.0" 形式
            if (value.Contains('.'))
            {
                if (!decimal.TryParse(value, NumberStyles.Number, CultureInfo.InvariantCulture, out decimal d) ||
                    d != Math.Truncate(d))
                {
                    yield return Issue(row, src, col, DvValidationLevel.Error, "整数格式错误", value,
                        "值不是整数格式");
                    yield break;
                }
                value = ((long)d).ToString();
            }
            if (!long.TryParse(value, out long iv))
            {
                yield return Issue(row, src, col, DvValidationLevel.Error, "整数格式错误", value, "值无法解析为整数");
                yield break;
            }
            var (min, max) = IntRange(col);
            if (iv < min || iv > max)
                yield return Issue(row, src, col, DvValidationLevel.Error, "整数溢出", value,
                    $"值 {iv} 超出范围 [{min}, {max}]");
        }

        private static (long min, long max) IntRange(DvTargetColumn col)
        {
            if (col.DatabaseType == DvDbType.SqlServer)
                return col.OriginalDataType switch
                {
                    "tinyint" => (0, 255),
                    "smallint" => (-32768, 32767),
                    _ => (int.MinValue, int.MaxValue)
                };
            return col.OriginalDataType switch
            {
                "smallint" or "int2" => (-32768, 32767),
                _ => (int.MinValue, int.MaxValue)
            };
        }

        // ── Long ──────────────────────────────────────────────
        private static IEnumerable<DvIssue> ValidateLong(string value, DvTargetColumn col, int row, string? src)
        {
            if (value.Contains('.'))
            {
                if (!decimal.TryParse(value, NumberStyles.Number, CultureInfo.InvariantCulture, out decimal d) ||
                    d != Math.Truncate(d))
                {
                    yield return Issue(row, src, col, DvValidationLevel.Error, "整数格式错误", value, "值不是整数格式");
                    yield break;
                }
            }
            if (!long.TryParse(value.Split('.')[0], out _))
                yield return Issue(row, src, col, DvValidationLevel.Error, "整数格式错误", value, "值无法解析为长整数");
        }

        // ── Decimal ───────────────────────────────────────────
        private static IEnumerable<DvIssue> ValidateDecimal(string value, DvTargetColumn col, int row, string? src)
        {
            if (!decimal.TryParse(value, NumberStyles.Number, CultureInfo.InvariantCulture, out decimal dv))
            {
                yield return Issue(row, src, col, DvValidationLevel.Error, "数值格式错误", value, "值无法解析为数值");
                yield break;
            }
            if (col.NumericPrecision.HasValue && col.NumericScale.HasValue)
            {
                int p = col.NumericPrecision.Value, s = col.NumericScale.Value;
                int maxIntDigits = p - s;

                string abs = Math.Abs(dv).ToString("G", CultureInfo.InvariantCulture);
                var parts = abs.Split('.');
                int intDigits = parts[0].TrimStart('0').Length;
                if (intDigits == 0) intDigits = 0;
                int scaleDigits = parts.Length > 1 ? parts[1].Length : 0;

                if (intDigits > maxIntDigits)
                    yield return Issue(row, src, col, DvValidationLevel.Error, "数值精度溢出", value,
                        $"整数部分 {intDigits} 位，超过限制 {maxIntDigits} 位 (precision={p},scale={s})");
                else if (scaleDigits > s)
                    yield return Issue(row, src, col, DvValidationLevel.Error, "小数位溢出", value,
                        $"小数部分 {scaleDigits} 位，超过限制 {s} 位");
            }
        }

        // ── Date ──────────────────────────────────────────────
        private static IEnumerable<DvIssue> ValidateDate(string value, DvTargetColumn col, int row, string? src)
        {
            if (!DateTime.TryParseExact(value, DateFormats, CultureInfo.InvariantCulture,
                DateTimeStyles.None, out _))
                yield return Issue(row, src, col, DvValidationLevel.Error, "日期格式错误", value,
                    $"日期格式不符合要求（支持: yyyy-MM-dd / yyyy/MM/dd / yyyyMMdd）");
        }

        // ── DateTime ──────────────────────────────────────────
        private static IEnumerable<DvIssue> ValidateDateTime(string value, DvTargetColumn col, int row, string? src)
        {
            if (!DateTime.TryParseExact(value, DateTimeFormats, CultureInfo.InvariantCulture,
                DateTimeStyles.None, out _))
                yield return Issue(row, src, col, DvValidationLevel.Error, "日期时间格式错误", value,
                    "日期时间格式不符合要求");
        }

        // ── Boolean ───────────────────────────────────────────
        private static readonly HashSet<string> BoolTrue = ["1", "true", "yes", "y", "是", "t"];
        private static readonly HashSet<string> BoolFalse = ["0", "false", "no", "n", "否", "f"];

        private static IEnumerable<DvIssue> ValidateBoolean(string value, DvTargetColumn col, int row, string? src)
        {
            var v = value.ToLowerInvariant();
            if (!BoolTrue.Contains(v) && !BoolFalse.Contains(v))
                yield return Issue(row, src, col, DvValidationLevel.Error, "布尔格式错误", value,
                    "布尔值格式无效（接受: 1/0/true/false/是/否/Y/N）");
        }

        // ── GUID ──────────────────────────────────────────────
        private static IEnumerable<DvIssue> ValidateGuid(string value, DvTargetColumn col, int row, string? src)
        {
            if (!Guid.TryParse(value, out _))
                yield return Issue(row, src, col, DvValidationLevel.Error, "GUID格式错误", value,
                    "GUID 格式无效");
        }

        // ── JSON ──────────────────────────────────────────────
        private static IEnumerable<DvIssue> ValidateJson(string value, DvTargetColumn col, int row, string? src)
        {
            bool valid = true;
            try { using var doc = System.Text.Json.JsonDocument.Parse(value); }
            catch { valid = false; }
            if (!valid)
                yield return Issue(row, src, col, DvValidationLevel.Error, "JSON格式错误", value, "JSON 格式无效");
        }

        private static DvIssue Issue(int row, string? src, DvTargetColumn col,
            DvValidationLevel level, string type, string? val, string msg) => new()
        {
            RowNumber = row,
            SourceColumnName = src,
            TargetColumnName = col.ColumnName,
            TargetDataType = col.DisplayType,
            Level = level,
            ErrorType = type,
            ActualValue = val?.Length > 200 ? val[..200] + "…" : val,
            Message = msg
        };
    }
}
