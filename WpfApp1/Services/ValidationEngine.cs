using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace WpfApp1.Services
{
    /// <summary>
    /// 数据验证核心引擎。
    /// 这层只关心三件事：按映射取值、按目标字段规则校验、产出统一的问题清单。
    /// </summary>
    // Core validation pipeline:
    // 1) resolve usable mappings for the current source batch,
    // 2) validate row by row against normalized target-column rules,
    // 3) emit a unified issue list for UI display and Excel export.
    public static class ValidationEngine
    {
        // 错误明细上限，避免异常数据一次性产出过多结果导致页面卡顿。
        private const int MaxErrors = 10000;

        // 仅日期字段允许的格式。
        private static readonly string[] DateFormats =
        [
            "yyyy-MM-dd",
            "yyyy/MM/dd",
            "yyyyMMdd"
        ];

        // 日期时间字段允许的格式。
        private static readonly string[] DateTimeFormats =
        [
            "yyyy-MM-dd HH:mm:ss",
            "yyyy/MM/dd HH:mm:ss",
            "yyyy-MM-dd HH:mm:ss.fff",
            "yyyy-MM-dd HH:mm:ss.ffffff",
            "yyyy-MM-dd HH:mm:ss.fffffff",
            "yyyy-MM-dd",
            "yyyy/MM/dd",
            "yyyyMMddHHmmss",
            "yyyyMMdd"
        ];

        // 纯时间字段允许的格式。
        private static readonly string[] TimeFormats =
        [
            "HH:mm",
            "HH:mm:ss",
            "HH:mm:ss.fff",
            "HH:mm:ss.ffffff",
            "HH:mm:ss.fffffff"
        ];

        // 带时区偏移的时间字段允许的格式。
        private static readonly string[] TimeWithOffsetFormats =
        [
            "HH:mmzzz",
            "HH:mm:sszzz",
            "HH:mm:ss.fffzzz",
            "HH:mm:ss.ffffffzzz",
            "HH:mm:ss.fffffffzzz"
        ];

        private static readonly HashSet<string> BoolTrue =
            ["1", "true", "yes", "y", "是", "t"];

        private static readonly HashSet<string> BoolFalse =
            ["0", "false", "no", "n", "否", "f"];

        /// <summary>
        /// 对当前映射后的整批源数据执行校验。
        /// </summary>
        // Batch entry point used by the validation page.
        // This method only validates fields that have a usable mapping, and it also
        // adds row-level errors for required target columns that are still unmapped.
        public static async Task<DvValidationResult> RunAsync(
            IReadOnlyList<DvTargetColumn> schema,
            DvSourceData data,
            IReadOnlyList<DvMappingRow> mappings,
            IReadOnlyList<string>? primaryKeyColumns = null,
            bool skipIntegerFormatErrors = false,
            bool skipGuidFormatErrors = false,
            bool skipDateTimeFormatErrors = false,
            IProgress<(int current, int total, int errors)>? progress = null,
            CancellationToken ct = default)
        {
            var start = DateTime.UtcNow;
            var issues = new List<DvIssue>();
            int processed = 0;
            int errorCount = 0;

            // 建立源表头索引，后续行循环里不再重复查找。
            var headerIndexMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < data.Headers.Count; i++)
            {
                headerIndexMap[data.Headers[i]] = i;
            }

            // 预解析主键列索引，便于结果里展示“主键=值”的定位信息。
            var primaryKeyIndices = new List<(string name, int index)>();
            if (primaryKeyColumns != null)
            {
                foreach (var primaryKey in primaryKeyColumns)
                {
                    if (headerIndexMap.TryGetValue(primaryKey, out int index))
                    {
                        primaryKeyIndices.Add((primaryKey, index));
                    }
                }
            }

            // 只保留真正可执行的“源字段映射”，避免每一行都去做映射判定。
            var resolvedMappings = new List<(DvTargetColumn column, int sourceIndex, string sourceName)>();
            var schemaMap = schema.ToDictionary(c => c.ColumnName, StringComparer.OrdinalIgnoreCase);
            foreach (var mapping in mappings)
            {
                if (mapping.MappingType != DvMappingType.Source || string.IsNullOrWhiteSpace(mapping.SourceColumnName))
                {
                    continue;
                }

                if (!schemaMap.TryGetValue(mapping.TargetColumnName, out var targetColumn))
                {
                    continue;
                }

                if (!headerIndexMap.TryGetValue(mapping.SourceColumnName, out int sourceIndex))
                {
                    continue;
                }

                resolvedMappings.Add((targetColumn, sourceIndex, mapping.SourceColumnName));
            }

            // 必填但没有可用映射的字段，整批数据的每一行都要报错一次。
            var unmappedRequiredColumns = new List<DvTargetColumn>();
            foreach (var column in schema.Where(c => !c.IsNullable))
            {
                var mapping = mappings.FirstOrDefault(m =>
                    string.Equals(m.TargetColumnName, column.ColumnName, StringComparison.OrdinalIgnoreCase));

                if (mapping == null || mapping.MappingType == DvMappingType.Ignore)
                {
                    // 自动生成候选字段（例如系统 UUID 主键）允许安全忽略。
                    if (mapping != null && mapping.IsAutoGenCandidate)
                    {
                        continue;
                    }

                    unmappedRequiredColumns.Add(column);
                }
            }

            int reportInterval = Math.Max(100, data.Rows.Count / 200);

            await Task.Run(() =>
            {
                for (int rowIndex = 0; rowIndex < data.Rows.Count; rowIndex++)
                {
                    ct.ThrowIfCancellationRequested();

                    var row = data.Rows[rowIndex];
                    int rowNumber = rowIndex + 2; // 第 1 行是表头，数据从第 2 行起。

                    string? primaryKeyDisplay = null;
                    if (primaryKeyIndices.Count > 0)
                    {
                        primaryKeyDisplay = string.Join(", ", primaryKeyIndices.Select(primaryKey =>
                            $"{primaryKey.name}={((primaryKey.index < row.Count ? row[primaryKey.index] : null) ?? "NULL")}"));
                    }

                    foreach (var (column, sourceIndex, sourceName) in resolvedMappings)
                    {
                        string? value = sourceIndex < row.Count ? row[sourceIndex] : null;

                        foreach (var issue in ValidateCell(
                            value,
                            column,
                            rowNumber,
                            sourceName,
                            skipIntegerFormatErrors,
                            skipGuidFormatErrors,
                            skipDateTimeFormatErrors))
                        {
                            issue.PrimaryKeyDisplay = primaryKeyDisplay;
                            if (issue.Level == DvValidationLevel.Error)
                            {
                                errorCount++;
                            }

                            issues.Add(issue);
                            if (issues.Count >= MaxErrors)
                            {
                                goto Done;
                            }
                        }
                    }

                    foreach (var requiredColumn in unmappedRequiredColumns)
                    {
                        errorCount++;
                        issues.Add(new DvIssue
                        {
                            RowNumber = rowNumber,
                            PrimaryKeyDisplay = primaryKeyDisplay,
                            TargetColumnName = requiredColumn.ColumnName,
                            TargetDataType = requiredColumn.DisplayType,
                            Level = DvValidationLevel.Error,
                            ErrorType = "必填未映射",
                            ActualValue = null,
                            Message = $"必填字段 {requiredColumn.ColumnName} 未映射或值为空"
                        });

                        if (issues.Count >= MaxErrors)
                        {
                            goto Done;
                        }
                    }

                    processed++;
                    if (processed % reportInterval == 0)
                    {
                        progress?.Report((processed, data.Rows.Count, errorCount));
                    }
                }

            Done:;
            }, ct);

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

        /// <summary>
        /// 单元格校验总入口。
        /// 先做空值判断，再按目标字段归一化类型分发到具体规则。
        /// </summary>
        // Single-cell validation dispatcher.
        // The order matters:
        // 1) required/empty checks,
        // 2) nullable empty short-circuit,
        // 3) type-specific rule dispatch.
        private static IEnumerable<DvIssue> ValidateCell(
            string? value,
            DvTargetColumn column,
            int rowNumber,
            string? sourceColumnName,
            bool skipIntegerFormatErrors = false,
            bool skipGuidFormatErrors = false,
            bool skipDateTimeFormatErrors = false)
        {
            bool isEmpty = string.IsNullOrWhiteSpace(value);
            if (!column.IsNullable && isEmpty)
            {
                yield return Issue(
                    rowNumber,
                    sourceColumnName,
                    column,
                    DvValidationLevel.Error,
                    "必填为空",
                    value,
                    "字段不可为空，实际值为空");
                yield break;
            }

            // 可空字段为空时，直接视为通过，不再继续做类型校验。
            if (isEmpty)
            {
                yield break;
            }

            string rawValue = value!;
            string normalizedValue = rawValue.Trim();

            switch (column.NormalizedType)
            {
                case DvNormalizedType.String:
                    foreach (var issue in ValidateString(rawValue, column, rowNumber, sourceColumnName))
                    {
                        yield return issue;
                    }
                    break;

                case DvNormalizedType.Integer:
                    if (!skipIntegerFormatErrors)
                    {
                        foreach (var issue in ValidateInteger(normalizedValue, column, rowNumber, sourceColumnName))
                        {
                            yield return issue;
                        }
                    }
                    break;

                case DvNormalizedType.Long:
                    if (!skipIntegerFormatErrors)
                    {
                        foreach (var issue in ValidateLong(normalizedValue, column, rowNumber, sourceColumnName))
                        {
                            yield return issue;
                        }
                    }
                    break;

                case DvNormalizedType.Decimal:
                    foreach (var issue in ValidateDecimal(normalizedValue, column, rowNumber, sourceColumnName))
                    {
                        yield return issue;
                    }
                    break;

                case DvNormalizedType.Date:
                    foreach (var issue in ValidateDate(normalizedValue, column, rowNumber, sourceColumnName))
                    {
                        yield return issue;
                    }
                    break;

                case DvNormalizedType.DateTime:
                    if (!skipDateTimeFormatErrors)
                    {
                        foreach (var issue in ValidateDateTime(normalizedValue, column, rowNumber, sourceColumnName))
                        {
                            yield return issue;
                        }
                    }
                    break;

                case DvNormalizedType.Time:
                    if (!skipDateTimeFormatErrors)
                    {
                        foreach (var issue in ValidateTime(normalizedValue, column, rowNumber, sourceColumnName))
                        {
                            yield return issue;
                        }
                    }
                    break;

                case DvNormalizedType.Boolean:
                    foreach (var issue in ValidateBoolean(normalizedValue, column, rowNumber, sourceColumnName))
                    {
                        yield return issue;
                    }
                    break;

                case DvNormalizedType.Guid:
                    if (!skipGuidFormatErrors)
                    {
                        foreach (var issue in ValidateGuid(normalizedValue, column, rowNumber, sourceColumnName))
                        {
                            yield return issue;
                        }
                    }
                    break;

                case DvNormalizedType.Json:
                    foreach (var issue in ValidateJson(normalizedValue, column, rowNumber, sourceColumnName))
                    {
                        yield return issue;
                    }
                    break;

                case DvNormalizedType.Unknown:
                    yield return Issue(
                        rowNumber,
                        sourceColumnName,
                        column,
                        DvValidationLevel.Warning,
                        "未识别类型",
                        rawValue,
                        $"字段类型 {column.OriginalDataType} 未识别，仅做非空校验");
                    break;
            }
        }

        /// <summary>
        /// 字符串规则：
        /// 1. 按原始值长度校验，保留前后空格参与长度判断。
        /// 2. SQL Server 的 varchar/char 若出现非 ASCII 字符，追加“字节可能超长”警告。
        /// </summary>
        // String rule set:
        // - enforce configured character length,
        // - keep leading/trailing spaces in the length calculation,
        // - emit a warning for SQL Server varchar/char when non-ASCII content may
        //   exceed the byte length even if the character count still fits.
        private static IEnumerable<DvIssue> ValidateString(string value, DvTargetColumn column, int rowNumber, string? sourceColumnName)
        {
            if (column.MaxLength.HasValue && column.MaxLength.Value > 0)
            {
                int length = value.Length;
                if (length > column.MaxLength.Value)
                {
                    yield return Issue(
                        rowNumber,
                        sourceColumnName,
                        column,
                        DvValidationLevel.Error,
                        "字符超长",
                        value,
                        $"字符长度 {length} 超过限制 {column.MaxLength}");
                }
                else if (column.DatabaseType == DvDbType.SqlServer &&
                         (string.Equals(column.OriginalDataType, "varchar", StringComparison.OrdinalIgnoreCase) ||
                          string.Equals(column.OriginalDataType, "char", StringComparison.OrdinalIgnoreCase)) &&
                         value.Any(c => c > 127))
                {
                    yield return Issue(
                        rowNumber,
                        sourceColumnName,
                        column,
                        DvValidationLevel.Warning,
                        "字节可能超长",
                        value,
                        "varchar/char 含多字节字符，实际字节占用可能超过限制");
                }
            }
        }

        /// <summary>
        /// 整数规则：
        /// 1. 允许 3.0 这类“数值上为整数”的输入。
        /// 2. 按数据库原始类型做范围校验。
        /// </summary>
        // Integer rule set:
        // - accept whole-number text and decimal text like 3.0,
        // - reject non-integral decimals,
        // - then enforce DB-specific integer range limits.
        private static IEnumerable<DvIssue> ValidateInteger(string value, DvTargetColumn column, int rowNumber, string? sourceColumnName)
        {
            if (value.Contains('.'))
            {
                if (!decimal.TryParse(value, NumberStyles.Number, CultureInfo.InvariantCulture, out decimal decimalValue) ||
                    decimalValue != Math.Truncate(decimalValue))
                {
                    yield return Issue(
                        rowNumber,
                        sourceColumnName,
                        column,
                        DvValidationLevel.Error,
                        "整数格式错误",
                        value,
                        "值不是整数格式");
                    yield break;
                }

                value = ((long)decimalValue).ToString(CultureInfo.InvariantCulture);
            }

            if (!long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out long integerValue))
            {
                yield return Issue(
                    rowNumber,
                    sourceColumnName,
                    column,
                    DvValidationLevel.Error,
                    "整数格式错误",
                    value,
                    "值无法解析为整数");
                yield break;
            }

            var (min, max) = IntRange(column);
            if (integerValue < min || integerValue > max)
            {
                yield return Issue(
                    rowNumber,
                    sourceColumnName,
                    column,
                    DvValidationLevel.Error,
                    "整数溢出",
                    value,
                    $"值 {integerValue} 超出范围 [{min}, {max}]");
            }
        }

        private static (long min, long max) IntRange(DvTargetColumn column)
        {
            if (column.DatabaseType == DvDbType.SqlServer)
            {
                return column.OriginalDataType switch
                {
                    "tinyint" => (0, 255),
                    "smallint" => (-32768, 32767),
                    _ => (int.MinValue, int.MaxValue)
                };
            }

            return column.OriginalDataType switch
            {
                "smallint" or "int2" => (-32768, 32767),
                _ => (int.MinValue, int.MaxValue)
            };
        }

        /// <summary>
        /// 长整数规则与整数类似，只是不再做人为的 int 范围限制。
        /// </summary>
        // Long rule set:
        // - same integral-shape check as integer,
        // - then rely on long parsing to enforce range.
        private static IEnumerable<DvIssue> ValidateLong(string value, DvTargetColumn column, int rowNumber, string? sourceColumnName)
        {
            if (value.Contains('.'))
            {
                if (!decimal.TryParse(value, NumberStyles.Number, CultureInfo.InvariantCulture, out decimal decimalValue) ||
                    decimalValue != Math.Truncate(decimalValue))
                {
                    yield return Issue(
                        rowNumber,
                        sourceColumnName,
                        column,
                        DvValidationLevel.Error,
                        "整数格式错误",
                        value,
                        "值不是整数格式");
                    yield break;
                }

                value = ((long)decimalValue).ToString(CultureInfo.InvariantCulture);
            }

            if (!long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out _))
            {
                yield return Issue(
                    rowNumber,
                    sourceColumnName,
                    column,
                    DvValidationLevel.Error,
                    "整数格式错误",
                    value,
                    "值无法解析为长整数");
            }
        }

        /// <summary>
        /// 数值规则：
        /// 1. 先判断是否可解析为十进制数。
        /// 2. 若结构中带 precision/scale，再分别校验整数位和小数位。
        /// </summary>
        // Decimal rule set:
        // - value must be parseable as an invariant-culture decimal,
        // - precision/scale are validated against the target definition when present.
        private static IEnumerable<DvIssue> ValidateDecimal(string value, DvTargetColumn column, int rowNumber, string? sourceColumnName)
        {
            if (!decimal.TryParse(value, NumberStyles.Number, CultureInfo.InvariantCulture, out decimal decimalValue))
            {
                yield return Issue(
                    rowNumber,
                    sourceColumnName,
                    column,
                    DvValidationLevel.Error,
                    "数值格式错误",
                    value,
                    "值无法解析为数值");
                yield break;
            }

            if (column.NumericPrecision.HasValue && column.NumericScale.HasValue)
            {
                int precision = column.NumericPrecision.Value;
                int scale = column.NumericScale.Value;
                int maxIntegerDigits = precision - scale;

                string absolute = Math.Abs(decimalValue).ToString("G", CultureInfo.InvariantCulture);
                var parts = absolute.Split('.');
                int integerDigits = parts[0].TrimStart('0').Length;
                int scaleDigits = parts.Length > 1 ? parts[1].Length : 0;

                if (integerDigits > maxIntegerDigits)
                {
                    yield return Issue(
                        rowNumber,
                        sourceColumnName,
                        column,
                        DvValidationLevel.Error,
                        "数值精度溢出",
                        value,
                        $"整数部分 {integerDigits} 位，超过限制 {maxIntegerDigits} 位 (precision={precision}, scale={scale})");
                }
                else if (scaleDigits > scale)
                {
                    yield return Issue(
                        rowNumber,
                        sourceColumnName,
                        column,
                        DvValidationLevel.Error,
                        "小数位溢出",
                        value,
                        $"小数部分 {scaleDigits} 位，超过限制 {scale} 位");
                }
            }
        }

        // Date rule set uses exact-match formats only.
        // This is intentionally stricter than DateTime.TryParse to keep import behavior predictable.
        private static IEnumerable<DvIssue> ValidateDate(string value, DvTargetColumn column, int rowNumber, string? sourceColumnName)
        {
            if (!DateTime.TryParseExact(
                    value,
                    DateFormats,
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.None,
                    out _))
            {
                yield return Issue(
                    rowNumber,
                    sourceColumnName,
                    column,
                    DvValidationLevel.Error,
                    "日期格式错误",
                    value,
                    "日期格式不符合要求（支持: yyyy-MM-dd / yyyy/MM/dd / yyyyMMdd）");
            }
        }

        // DateTime rule set allows the accepted full timestamp formats and a small
        // set of date-only fallbacks used by the current import workflow.
        private static IEnumerable<DvIssue> ValidateDateTime(string value, DvTargetColumn column, int rowNumber, string? sourceColumnName)
        {
            if (!DateTime.TryParseExact(
                    value,
                    DateTimeFormats,
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.None,
                    out _))
            {
                yield return Issue(
                    rowNumber,
                    sourceColumnName,
                    column,
                    DvValidationLevel.Error,
                    "日期时间格式错误",
                    value,
                    "日期时间格式不符合要求");
            }
        }

        // Time rule set supports pure time values and time-with-offset values.
        private static IEnumerable<DvIssue> ValidateTime(string value, DvTargetColumn column, int rowNumber, string? sourceColumnName)
        {
            bool isTimeValid =
                DateTime.TryParseExact(value, TimeFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out _) ||
                DateTimeOffset.TryParseExact(value, TimeWithOffsetFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out _);

            if (!isTimeValid)
            {
                yield return Issue(
                    rowNumber,
                    sourceColumnName,
                    column,
                    DvValidationLevel.Error,
                    "时间格式错误",
                    value,
                    "时间格式不符合要求");
            }
        }

        // Boolean rule set uses an explicit allow-list instead of general parsing so
        // the accepted source representations stay stable across runtimes.
        private static IEnumerable<DvIssue> ValidateBoolean(string value, DvTargetColumn column, int rowNumber, string? sourceColumnName)
        {
            var normalized = value.ToLowerInvariant();
            if (!BoolTrue.Contains(normalized) && !BoolFalse.Contains(normalized))
            {
                yield return Issue(
                    rowNumber,
                    sourceColumnName,
                    column,
                    DvValidationLevel.Error,
                    "布尔格式错误",
                    value,
                    "布尔值格式无效（接受: 1/0/true/false/是/否/Y/N）");
            }
        }

        // GUID / UUID rule set.
        private static IEnumerable<DvIssue> ValidateGuid(string value, DvTargetColumn column, int rowNumber, string? sourceColumnName)
        {
            if (!Guid.TryParse(value, out _))
            {
                yield return Issue(
                    rowNumber,
                    sourceColumnName,
                    column,
                    DvValidationLevel.Error,
                    "GUID格式错误",
                    value,
                    "GUID 格式无效");
            }
        }

        // JSON rule set only verifies syntax. It does not validate business schema.
        private static IEnumerable<DvIssue> ValidateJson(string value, DvTargetColumn column, int rowNumber, string? sourceColumnName)
        {
            var isValidJson = true;

            try
            {
                using var document = JsonDocument.Parse(value);
            }
            catch
            {
                isValidJson = false;
            }

            if (!isValidJson)
            {
                yield return Issue(
                    rowNumber,
                    sourceColumnName,
                    column,
                    DvValidationLevel.Error,
                    "JSON格式错误",
                    value,
                    "JSON 格式无效");
            }
        }

        // Unified issue factory used by every rule branch so UI/export receive a
        // consistent payload shape. Long actual values are trimmed here for display safety.
        private static DvIssue Issue(
            int rowNumber,
            string? sourceColumnName,
            DvTargetColumn column,
            DvValidationLevel level,
            string errorType,
            string? actualValue,
            string message) => new()
        {
            RowNumber = rowNumber,
            SourceColumnName = sourceColumnName,
            TargetColumnName = column.ColumnName,
            TargetDataType = column.DisplayType,
            Level = level,
            ErrorType = errorType,
            ActualValue = actualValue?.Length > 200 ? actualValue[..200] + "…" : actualValue,
            Message = message
        };
    }
}
