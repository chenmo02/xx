using System.Data;
using System.Globalization;
using System.Text;

namespace WpfApp1.Services
{
    public static class SqlGeneratorService
    {
        private enum ColumnKind
        {
            String,
            Integer,
            Decimal,
            Boolean,
            DateTime
        }

        private sealed class ColumnDefinition
        {
            public required string OriginalName { get; init; }

            public required string SafeName { get; init; }

            public required ColumnKind Kind { get; init; }

            public required string SqlType { get; init; }
        }

        public enum DbType
        {
            PostgreSQL,
            SqlServer,
            MySQL,
            Oracle
        }

        public static string GetDefaultTableName(DbType dbType) => dbType switch
        {
            DbType.PostgreSQL => "TempTable",
            DbType.MySQL => "temp_table",
            DbType.Oracle => "TEMP_TABLE",
            _ => "#TMP"
        };

        public static string GetDefaultTableName(DbType dbType, string? prefix)
        {
            if (string.IsNullOrWhiteSpace(prefix))
            {
                return GetDefaultTableName(dbType);
            }

            string normalizedPrefix = NormalizeTablePrefix(prefix);
            if (string.IsNullOrWhiteSpace(normalizedPrefix))
            {
                return GetDefaultTableName(dbType);
            }

            string logicalName = $"{normalizedPrefix}_data";
            return dbType switch
            {
                DbType.SqlServer => $"#{logicalName}",
                DbType.Oracle => logicalName.ToUpperInvariant(),
                _ => logicalName
            };
        }

        public static string NormalizeTableName(DbType dbType, string? tableName, string? prefix = null)
        {
            string candidate = string.IsNullOrWhiteSpace(tableName)
                ? GetDefaultTableName(dbType, prefix)
                : tableName.Trim();

            if (string.IsNullOrWhiteSpace(candidate))
            {
                return GetDefaultTableName(dbType, prefix);
            }

            if (dbType == DbType.SqlServer && candidate.StartsWith("#", StringComparison.Ordinal))
            {
                return candidate;
            }

            if (TryGetTemporaryTableLogicalName(candidate, out string logicalName))
            {
                return dbType switch
                {
                    DbType.SqlServer => $"#{logicalName}",
                    DbType.Oracle => logicalName.ToUpperInvariant(),
                    DbType.MySQL => logicalName,
                    _ => logicalName
                };
            }

            if (dbType == DbType.SqlServer && TryGetSqlServerSimpleTableName(candidate, out string sqlServerTableName))
            {
                return $"#{sqlServerTableName}";
            }

            return candidate;
        }

        public static string GenerateFullSql(
            DbType dbType,
            string tableName,
            DataTable data,
            bool dropIfExists = true,
            bool batchInsert = true,
            int batchSize = 1000,
            bool limitStringLength = true)
        {
            var columns = BuildColumnDefinitions(dbType, data, limitStringLength);
            var sb = new StringBuilder();

            sb.AppendLine("-- ============================================");
            sb.AppendLine("-- 自动生成的数据导入 SQL");
            sb.AppendLine($"-- 数据库类型: {dbType}");
            sb.AppendLine($"-- 表名: {tableName}");
            sb.AppendLine($"-- 列数: {data.Columns.Count}");
            sb.AppendLine($"-- 数据行数: {data.Rows.Count}");
            sb.AppendLine($"-- 生成时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            sb.AppendLine("-- ============================================");
            sb.AppendLine();

            sb.Append(GenerateCreateTableSql(dbType, tableName, columns, dropIfExists));
            sb.AppendLine();

            if (data.Rows.Count > 0)
            {
                sb.Append(GenerateInsertSql(dbType, tableName, data, columns, batchInsert, batchSize));
                sb.AppendLine();
            }

            sb.AppendLine("-- 验证导入结果");
            sb.AppendLine($"SELECT COUNT(*) AS total_rows FROM {WrapName(dbType, tableName)};");
            sb.AppendLine($"-- SELECT * FROM {WrapName(dbType, tableName)}{GetPreviewSuffix(dbType)};");

            return sb.ToString();
        }

        public static string GenerateCreateTableSql(
            DbType dbType,
            string tableName,
            DataTable data,
            bool dropIfExists,
            bool limitStringLength = true)
        {
            var columns = BuildColumnDefinitions(dbType, data, limitStringLength);
            return GenerateCreateTableSql(dbType, tableName, columns, dropIfExists);
        }

        public static string GenerateInsertSql(
            DbType dbType,
            string tableName,
            DataTable data,
            bool batchInsert,
            int batchSize,
            bool limitStringLength = true)
        {
            var columns = BuildColumnDefinitions(dbType, data, limitStringLength);
            return GenerateInsertSql(dbType, tableName, data, columns, batchInsert, batchSize);
        }

        private static string GenerateCreateTableSql(
            DbType dbType,
            string tableName,
            IReadOnlyList<ColumnDefinition> columns,
            bool dropIfExists)
        {
            var sb = new StringBuilder();
            string wrappedName = WrapName(dbType, tableName);

            if (dropIfExists)
            {
                switch (dbType)
                {
                    case DbType.PostgreSQL:
                    case DbType.MySQL:
                        sb.AppendLine($"DROP TABLE IF EXISTS {wrappedName};");
                        break;
                    case DbType.Oracle:
                        sb.AppendLine("BEGIN");
                        sb.AppendLine($"    EXECUTE IMMEDIATE 'DROP TABLE {tableName}';");
                        sb.AppendLine("EXCEPTION");
                        sb.AppendLine("    WHEN OTHERS THEN");
                        sb.AppendLine("        IF SQLCODE != -942 THEN RAISE; END IF;");
                        sb.AppendLine("END;");
                        sb.AppendLine("/");
                        break;
                    case DbType.SqlServer:
                        if (tableName.StartsWith("#", StringComparison.Ordinal))
                        {
                            sb.AppendLine($"IF OBJECT_ID('tempdb..{tableName}') IS NOT NULL DROP TABLE {tableName};");
                        }
                        else
                        {
                            sb.AppendLine($"IF OBJECT_ID('{tableName}', 'U') IS NOT NULL DROP TABLE {wrappedName};");
                        }
                        break;
                }

                sb.AppendLine();
            }

            sb.AppendLine(GetCreateTablePrefix(dbType, wrappedName));
            for (int i = 0; i < columns.Count; i++)
            {
                ColumnDefinition column = columns[i];
                string wrappedColumn = WrapName(dbType, column.SafeName);
                sb.Append($"    {wrappedColumn} {column.SqlType} NULL");
                sb.AppendLine(i < columns.Count - 1 ? "," : string.Empty);
            }

            sb.AppendLine(GetCreateTableSuffix(dbType));
            return sb.ToString();
        }

        private static string GenerateInsertSql(
            DbType dbType,
            string tableName,
            DataTable data,
            IReadOnlyList<ColumnDefinition> columns,
            bool batchInsert,
            int batchSize)
        {
            var sb = new StringBuilder();
            string wrappedTable = WrapName(dbType, tableName);
            string columnList = string.Join(", ", columns.Select(column => WrapName(dbType, column.SafeName)));

            sb.AppendLine($"-- 插入数据 ({data.Rows.Count} 行)");

            if (!batchInsert)
            {
                foreach (DataRow row in data.Rows)
                {
                    sb.Append($"INSERT INTO {wrappedTable} ({columnList}) VALUES (");
                    sb.Append(string.Join(", ", columns.Select((column, index) => FormatValue(row[index], column.Kind, dbType))));
                    sb.AppendLine(");");
                }

                return sb.ToString();
            }

            int effectiveBatchSize = dbType == DbType.SqlServer ? Math.Min(Math.Max(batchSize, 1), 1000) : Math.Max(batchSize, 1);

            if (dbType == DbType.Oracle)
            {
                for (int start = 0; start < data.Rows.Count; start += effectiveBatchSize)
                {
                    int end = Math.Min(start + effectiveBatchSize, data.Rows.Count);
                    sb.AppendLine("INSERT ALL");
                    for (int rowIndex = start; rowIndex < end; rowIndex++)
                    {
                        DataRow row = data.Rows[rowIndex];
                        string values = string.Join(", ", columns.Select((column, index) => FormatValue(row[index], column.Kind, dbType)));
                        sb.AppendLine($"    INTO {wrappedTable} ({columnList}) VALUES ({values})");
                    }

                    sb.AppendLine("SELECT 1 FROM DUAL;");
                    sb.AppendLine();
                }

                return sb.ToString();
            }

            for (int start = 0; start < data.Rows.Count; start += effectiveBatchSize)
            {
                int end = Math.Min(start + effectiveBatchSize, data.Rows.Count);
                sb.AppendLine($"INSERT INTO {wrappedTable} ({columnList}) VALUES");

                for (int rowIndex = start; rowIndex < end; rowIndex++)
                {
                    DataRow row = data.Rows[rowIndex];
                    string values = string.Join(", ", columns.Select((column, index) => FormatValue(row[index], column.Kind, dbType)));
                    sb.Append($"    ({values})");
                    sb.AppendLine(rowIndex < end - 1 ? "," : ";");
                }

                sb.AppendLine();
            }

            return sb.ToString();
        }

        private static List<ColumnDefinition> BuildColumnDefinitions(DbType dbType, DataTable data, bool limitStringLength)
        {
            var columns = new List<ColumnDefinition>(data.Columns.Count);
            var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            for (int index = 0; index < data.Columns.Count; index++)
            {
                DataColumn column = data.Columns[index];
                string safeName = MakeUniqueColumnName(SanitizeColumnName(column.ColumnName), usedNames);
                // Keep imported values as-is in generated SQL instead of normalizing booleans or numerics.
                ColumnKind kind = ColumnKind.String;
                int maxLength = GetMaxStringLength(data, index);
                string sqlType = GetSqlType(dbType, kind, maxLength, limitStringLength);

                columns.Add(new ColumnDefinition
                {
                    OriginalName = column.ColumnName,
                    SafeName = safeName,
                    Kind = kind,
                    SqlType = sqlType
                });
            }

            return columns;
        }

        private static ColumnKind InferColumnKind(DataColumn column, DataTable data, int columnIndex)
        {
            Type dataType = Nullable.GetUnderlyingType(column.DataType) ?? column.DataType;
            if (dataType == typeof(bool))
            {
                return ColumnKind.Boolean;
            }

            if (dataType == typeof(DateTime))
            {
                return ColumnKind.DateTime;
            }

            if (dataType == typeof(short) || dataType == typeof(int) || dataType == typeof(long))
            {
                return ColumnKind.Integer;
            }

            if (dataType == typeof(decimal) || dataType == typeof(double) || dataType == typeof(float))
            {
                return ColumnKind.Decimal;
            }

            bool hasValue = false;
            bool allBool = true;
            bool allInt = true;
            bool allDecimal = true;
            bool allDate = true;

            int sampleCount = Math.Min(200, data.Rows.Count);
            for (int rowIndex = 0; rowIndex < sampleCount; rowIndex++)
            {
                object raw = data.Rows[rowIndex][columnIndex];
                if (raw == DBNull.Value)
                {
                    continue;
                }

                string value = raw.ToString()?.Trim() ?? "";
                if (string.IsNullOrEmpty(value))
                {
                    continue;
                }

                hasValue = true;
                if (allBool && !TryParseBoolean(value, out _))
                {
                    allBool = false;
                }

                if (allInt && !long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out _) &&
                    !long.TryParse(value, NumberStyles.Integer, CultureInfo.CurrentCulture, out _))
                {
                    allInt = false;
                }

                if (allDecimal && !TryParseDecimal(value, out _))
                {
                    allDecimal = false;
                }

                if (allDate && !DateTime.TryParse(value, CultureInfo.CurrentCulture, DateTimeStyles.None, out _) &&
                    !DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.None, out _))
                {
                    allDate = false;
                }
            }

            if (!hasValue)
            {
                return ColumnKind.String;
            }

            if (allBool)
            {
                return ColumnKind.Boolean;
            }

            if (allInt)
            {
                return ColumnKind.Integer;
            }

            if (allDecimal)
            {
                return ColumnKind.Decimal;
            }

            if (allDate)
            {
                return ColumnKind.DateTime;
            }

            return ColumnKind.String;
        }

        private static int GetMaxStringLength(DataTable data, int columnIndex)
        {
            int maxLength = 0;
            foreach (DataRow row in data.Rows)
            {
                object raw = row[columnIndex];
                if (raw == DBNull.Value)
                {
                    continue;
                }

                maxLength = Math.Max(maxLength, raw.ToString()?.Length ?? 0);
            }

            return maxLength;
        }

        private static string GetSqlType(DbType dbType, ColumnKind kind, int maxLength, bool limitStringLength)
        {
            return kind switch
            {
                ColumnKind.Boolean => dbType switch
                {
                    DbType.PostgreSQL => "BOOLEAN",
                    DbType.MySQL => "TINYINT(1)",
                    DbType.Oracle => "NUMBER(1)",
                    _ => "BIT"
                },
                ColumnKind.Integer => dbType switch
                {
                    DbType.PostgreSQL => "BIGINT",
                    DbType.MySQL => "BIGINT",
                    DbType.Oracle => "NUMBER(19)",
                    _ => "BIGINT"
                },
                ColumnKind.Decimal => dbType switch
                {
                    DbType.PostgreSQL => "NUMERIC(18,6)",
                    DbType.MySQL => "DECIMAL(18,6)",
                    DbType.Oracle => "NUMBER(18,6)",
                    _ => "DECIMAL(18,6)"
                },
                ColumnKind.DateTime => dbType switch
                {
                    DbType.PostgreSQL => "TIMESTAMP",
                    DbType.MySQL => "DATETIME",
                    DbType.Oracle => "DATE",
                    _ => "DATETIME"
                },
                _ => GetStringSqlType(dbType, maxLength, limitStringLength)
            };
        }

        private static string GetStringSqlType(DbType dbType, int maxLength, bool limitStringLength)
        {
            int effectiveLength = limitStringLength
                ? 1000
                : Math.Max(1, Math.Min(Math.Max(maxLength, 1), 4000));

            if (!limitStringLength && maxLength > 4000)
            {
                return dbType switch
                {
                    DbType.PostgreSQL => "TEXT",
                    DbType.MySQL => "LONGTEXT",
                    DbType.Oracle => "CLOB",
                    _ => "NVARCHAR(MAX)"
                };
            }

            return dbType switch
            {
                DbType.PostgreSQL => $"VARCHAR({effectiveLength})",
                DbType.MySQL => $"VARCHAR({effectiveLength})",
                DbType.Oracle => $"VARCHAR2({effectiveLength})",
                _ => $"NVARCHAR({effectiveLength})"
            };
        }

        private static string FormatValue(object value, ColumnKind kind, DbType dbType)
        {
            if (value == DBNull.Value)
            {
                return "NULL";
            }

            string text = value.ToString()?.Trim() ?? "";
            if (string.IsNullOrEmpty(text))
            {
                return "NULL";
            }

            return kind switch
            {
                ColumnKind.Boolean => FormatBooleanValue(text, dbType),
                ColumnKind.Integer => FormatIntegerValue(text),
                ColumnKind.Decimal => FormatDecimalValue(text),
                ColumnKind.DateTime => FormatDateTimeValue(value, text, dbType),
                _ => $"'{EscapeString(text)}'"
            };
        }

        private static string FormatBooleanValue(string value, DbType dbType)
        {
            bool parsed = TryParseBoolean(value, out bool boolValue) && boolValue;
            return dbType switch
            {
                DbType.PostgreSQL => parsed ? "TRUE" : "FALSE",
                _ => parsed ? "1" : "0"
            };
        }

        private static string FormatIntegerValue(string value)
        {
            if (long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out long number) ||
                long.TryParse(value, NumberStyles.Integer, CultureInfo.CurrentCulture, out number))
            {
                return number.ToString(CultureInfo.InvariantCulture);
            }

            return "NULL";
        }

        private static string FormatDecimalValue(string value)
        {
            if (TryParseDecimal(value, out decimal number))
            {
                return number.ToString(CultureInfo.InvariantCulture);
            }

            return "NULL";
        }

        private static string FormatDateTimeValue(object rawValue, string value, DbType dbType)
        {
            DateTime dateTime;
            if (rawValue is DateTime exactDate)
            {
                dateTime = exactDate;
            }
            else if (!DateTime.TryParse(value, CultureInfo.CurrentCulture, DateTimeStyles.None, out dateTime) &&
                     !DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.None, out dateTime))
            {
                return "NULL";
            }

            return dbType switch
            {
                DbType.Oracle => $"TO_DATE('{dateTime:yyyy-MM-dd HH:mm:ss}', 'YYYY-MM-DD HH24:MI:SS')",
                _ => $"'{dateTime:yyyy-MM-dd HH:mm:ss}'"
            };
        }

        private static bool TryParseBoolean(string value, out bool result)
        {
            string normalized = value.Trim().ToLowerInvariant();
            switch (normalized)
            {
                case "true":
                case "t":
                case "yes":
                case "y":
                    result = true;
                    return true;
                case "false":
                case "f":
                case "no":
                case "n":
                    result = false;
                    return true;
                default:
                    return bool.TryParse(value, out result);
            }
        }

        private static bool TryParseDecimal(string value, out decimal result)
        {
            return decimal.TryParse(value, NumberStyles.Number, CultureInfo.InvariantCulture, out result) ||
                   decimal.TryParse(value, NumberStyles.Number, CultureInfo.CurrentCulture, out result);
        }

        private static string GetCreateTablePrefix(DbType dbType, string wrappedName) => dbType switch
        {
            DbType.PostgreSQL => $"CREATE TEMPORARY TABLE {wrappedName} (",
            DbType.MySQL => $"CREATE TEMPORARY TABLE {wrappedName} (",
            DbType.Oracle => $"CREATE GLOBAL TEMPORARY TABLE {wrappedName} (",
            _ => $"CREATE TABLE {wrappedName} ("
        };

        private static string GetCreateTableSuffix(DbType dbType) => dbType switch
        {
            DbType.Oracle => ") ON COMMIT PRESERVE ROWS;",
            _ => ");"
        };

        private static string GetPreviewSuffix(DbType dbType) => dbType switch
        {
            DbType.SqlServer => "",
            DbType.Oracle => " FETCH FIRST 10 ROWS ONLY",
            _ => " LIMIT 10"
        };

        private static string WrapName(DbType dbType, string name) => dbType switch
        {
            DbType.PostgreSQL => $"\"{name}\"",
            DbType.MySQL => $"`{name}`",
            DbType.Oracle => $"\"{name}\"",
            _ => name.StartsWith("#", StringComparison.Ordinal) ? name : $"[{name}]"
        };

        private static string SanitizeColumnName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return "col";
            }

            var sb = new StringBuilder();
            foreach (char c in name.Trim())
            {
                if (char.IsLetterOrDigit(c) || c == '_' || c > 127)
                {
                    sb.Append(c);
                }
                else
                {
                    sb.Append('_');
                }
            }

            string result = sb.ToString().Trim('_');
            if (string.IsNullOrWhiteSpace(result))
            {
                result = "col";
            }

            if (char.IsDigit(result[0]))
            {
                result = $"c_{result}";
            }

            return result;
        }

        private static string MakeUniqueColumnName(string name, ISet<string> usedNames)
        {
            string uniqueName = name;
            int suffix = 1;
            while (!usedNames.Add(uniqueName))
            {
                suffix++;
                uniqueName = $"{name}_{suffix}";
            }

            return uniqueName;
        }

        private static bool TryGetTemporaryTableLogicalName(string tableName, out string logicalName)
        {
            logicalName = tableName.Trim();
            if (string.IsNullOrWhiteSpace(logicalName))
            {
                logicalName = string.Empty;
                return false;
            }

            logicalName = StripIdentifierDelimiters(logicalName);
            logicalName = logicalName.TrimStart('#');

            if (string.IsNullOrWhiteSpace(logicalName) || logicalName.Contains('.'))
            {
                logicalName = string.Empty;
                return false;
            }

            string normalized = logicalName.Replace("_", string.Empty, StringComparison.Ordinal).ToLowerInvariant();
            return normalized is "temptable" or "tmp";
        }

        private static bool TryGetSqlServerSimpleTableName(string tableName, out string logicalName)
        {
            logicalName = StripIdentifierDelimiters(tableName.Trim());
            if (string.IsNullOrWhiteSpace(logicalName) || logicalName.Contains('.'))
            {
                logicalName = string.Empty;
                return false;
            }

            foreach (char c in logicalName)
            {
                if (!(char.IsLetterOrDigit(c) || c == '_' || c > 127))
                {
                    logicalName = string.Empty;
                    return false;
                }
            }

            return true;
        }

        private static string StripIdentifierDelimiters(string value)
        {
            if (value.Length >= 2)
            {
                char first = value[0];
                char last = value[^1];
                if ((first == '[' && last == ']') ||
                    (first == '"' && last == '"') ||
                    (first == '`' && last == '`'))
                {
                    return value[1..^1].Trim();
                }
            }

            return value;
        }

        private static string NormalizeTablePrefix(string prefix)
        {
            var sb = new StringBuilder();
            foreach (char c in prefix.Trim())
            {
                if (char.IsLetterOrDigit(c) || c == '_')
                {
                    sb.Append(c);
                }
                else
                {
                    sb.Append('_');
                }
            }

            return sb.ToString().Trim('_');
        }

        private static string EscapeString(string value) => value.Replace("'", "''");
    }
}
