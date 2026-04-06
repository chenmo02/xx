using System.Data;
using System.Text;

namespace WpfApp1.Services
{
    /// <summary>
    /// SQL 语句生成服务：根据 DataTable 生成建表和插入语句
    /// </summary>
    public static class SqlGeneratorService
    {
        public enum DbType
        {
            PostgreSQL,
            SqlServer
        }

        /// <summary>
        /// 生成完整 SQL（建表 + 插入）
        /// </summary>
        public static string GenerateFullSql(
            DbType dbType, string tableName, DataTable data,
            bool dropIfExists = true, bool batchInsert = true, int batchSize = 100)
        {
            var sb = new StringBuilder();

            // 头部注释
            sb.AppendLine($"-- ============================================");
            sb.AppendLine($"-- 自动生成的临时表 SQL");
            sb.AppendLine($"-- 数据库类型: {dbType}");
            sb.AppendLine($"-- 表名: {tableName}");
            sb.AppendLine($"-- 列数: {data.Columns.Count}");
            sb.AppendLine($"-- 数据行数: {data.Rows.Count}");
            sb.AppendLine($"-- 生成时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            sb.AppendLine($"-- ============================================");
            sb.AppendLine();

            // 建表语句
            sb.Append(GenerateCreateTableSql(dbType, tableName, data, dropIfExists));
            sb.AppendLine();

            // 插入语句
            if (data.Rows.Count > 0)
            {
                sb.Append(GenerateInsertSql(dbType, tableName, data, batchInsert, batchSize));
                sb.AppendLine();

                // 查询验证
                sb.AppendLine($"-- 验证导入结果");
                if (dbType == DbType.PostgreSQL)
                    sb.AppendLine($"SELECT COUNT(*) AS total_rows FROM \"{tableName}\";");
                else
                    sb.AppendLine($"SELECT COUNT(*) AS total_rows FROM [{tableName}];");

                sb.AppendLine($"-- SELECT * FROM {WrapName(dbType, tableName)} LIMIT 10;");
            }

            return sb.ToString();
        }

        /// <summary>
        /// 生成 CREATE TABLE 语句
        /// </summary>
        public static string GenerateCreateTableSql(DbType dbType, string tableName, DataTable data, bool dropIfExists)
        {
            var sb = new StringBuilder();
            string wrappedName = WrapName(dbType, tableName);

            // DROP IF EXISTS
            if (dropIfExists)
            {
                if (dbType == DbType.PostgreSQL)
                {
                    sb.AppendLine($"DROP TABLE IF EXISTS {wrappedName};");
                }
                else
                {
                    // SQL Server 临时表在 tempdb 中
                    if (tableName.StartsWith("#"))
                        sb.AppendLine($"IF OBJECT_ID('tempdb..{tableName}') IS NOT NULL DROP TABLE {tableName};");
                    else
                        sb.AppendLine($"IF OBJECT_ID('{tableName}', 'U') IS NOT NULL DROP TABLE {wrappedName};");
                }
                sb.AppendLine();
            }

            // CREATE TABLE
            if (dbType == DbType.PostgreSQL)
            {
                sb.AppendLine($"CREATE TEMP TABLE {wrappedName} (");
            }
            else
            {
                sb.AppendLine($"CREATE TABLE {wrappedName} (");
            }

            for (int i = 0; i < data.Columns.Count; i++)
            {
                string colName = SanitizeColumnName(data.Columns[i].ColumnName);
                string colType = InferColumnType(dbType, data, i);
                string wrappedCol = WrapName(dbType, colName);

                sb.Append($"    {wrappedCol} {colType}");
                if (i < data.Columns.Count - 1)
                    sb.AppendLine(",");
                else
                    sb.AppendLine();
            }

            sb.AppendLine(");");
            return sb.ToString();
        }

        /// <summary>
        /// 生成 INSERT 语句
        /// </summary>
        public static string GenerateInsertSql(DbType dbType, string tableName, DataTable data, bool batchInsert, int batchSize)
        {
            var sb = new StringBuilder();
            string wrappedName = WrapName(dbType, tableName);

            // 构建列名列表
            var colNames = new List<string>();
            for (int i = 0; i < data.Columns.Count; i++)
            {
                string colName = SanitizeColumnName(data.Columns[i].ColumnName);
                colNames.Add(WrapName(dbType, colName));
            }
            string colList = string.Join(", ", colNames);

            sb.AppendLine($"-- 插入数据 ({data.Rows.Count} 行)");

            if (batchInsert)
            {
                // 批量 INSERT（多行 VALUES）
                // SQL Server 单条 INSERT 最多 1000 行
                int effectiveBatch = dbType == DbType.SqlServer
                    ? Math.Min(batchSize, 1000)
                    : batchSize;

                for (int i = 0; i < data.Rows.Count; i += effectiveBatch)
                {
                    int end = Math.Min(i + effectiveBatch, data.Rows.Count);

                    sb.AppendLine($"INSERT INTO {wrappedName} ({colList}) VALUES");

                    for (int r = i; r < end; r++)
                    {
                        sb.Append("    (");
                        for (int c = 0; c < data.Columns.Count; c++)
                        {
                            sb.Append(EscapeValue(data.Rows[r][c]?.ToString()));
                            if (c < data.Columns.Count - 1)
                                sb.Append(", ");
                        }
                        sb.Append(')');

                        if (r < end - 1)
                            sb.AppendLine(",");
                        else
                            sb.AppendLine(";");
                    }
                    sb.AppendLine();
                }
            }
            else
            {
                // 逐行 INSERT
                foreach (DataRow row in data.Rows)
                {
                    sb.Append($"INSERT INTO {wrappedName} ({colList}) VALUES (");
                    for (int c = 0; c < data.Columns.Count; c++)
                    {
                        sb.Append(EscapeValue(row[c]?.ToString()));
                        if (c < data.Columns.Count - 1)
                            sb.Append(", ");
                    }
                    sb.AppendLine(");");
                }
                sb.AppendLine();
            }

            return sb.ToString();
        }

        /// <summary>
        /// 根据数据内容推断列类型
        /// </summary>
        private static string InferColumnType(DbType dbType, DataTable data, int colIndex)
        {
            // 采样前200行来推断类型
            int sampleCount = Math.Min(200, data.Rows.Count);
            bool allEmpty = true;
            bool allInt = true;
            bool allDecimal = true;
            bool allDate = true;
            int maxLen = 0;

            for (int r = 0; r < sampleCount; r++)
            {
                string val = data.Rows[r][colIndex]?.ToString()?.Trim() ?? "";
                if (string.IsNullOrEmpty(val)) continue;

                allEmpty = false;
                maxLen = Math.Max(maxLen, val.Length);

                if (allInt && !long.TryParse(val, out _))
                    allInt = false;

                if (allDecimal && !decimal.TryParse(val, out _))
                    allDecimal = false;

                if (allDate && !DateTime.TryParse(val, out _))
                    allDate = false;
            }

            // 如果全部为空或数据量为 0，默认 TEXT
            if (allEmpty || data.Rows.Count == 0)
            {
                return dbType == DbType.PostgreSQL ? "TEXT" : "NVARCHAR(MAX)";
            }

            // 推断类型
            if (allInt)
            {
                return dbType == DbType.PostgreSQL ? "BIGINT" : "BIGINT";
            }

            if (allDecimal)
            {
                return dbType == DbType.PostgreSQL ? "NUMERIC" : "DECIMAL(18,6)";
            }

            if (allDate)
            {
                return dbType == DbType.PostgreSQL ? "TIMESTAMP" : "DATETIME";
            }

            // 默认文本
            if (dbType == DbType.PostgreSQL)
            {
                return "TEXT";
            }
            else
            {
                if (maxLen <= 50) return "NVARCHAR(100)";
                if (maxLen <= 200) return "NVARCHAR(500)";
                if (maxLen <= 2000) return "NVARCHAR(4000)";
                return "NVARCHAR(MAX)";
            }
        }

        /// <summary>
        /// 转义 SQL 值
        /// </summary>
        private static string EscapeValue(string? value)
        {
            if (string.IsNullOrEmpty(value))
                return "NULL";

            // 转义单引号
            string escaped = value.Replace("'", "''");
            return $"'{escaped}'";
        }

        /// <summary>
        /// 包裹标识符名称
        /// </summary>
        private static string WrapName(DbType dbType, string name)
        {
            if (dbType == DbType.PostgreSQL)
                return $"\"{name}\"";
            else
            {
                // SQL Server 临时表 # 开头不加方括号
                if (name.StartsWith("#"))
                    return name;
                return $"[{name}]";
            }
        }

        /// <summary>
        /// 清理列名
        /// </summary>
        private static string SanitizeColumnName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return "unnamed";

            var sb = new StringBuilder();
            foreach (char c in name.Trim())
            {
                if (char.IsLetterOrDigit(c) || c == '_' || c > 127)
                    sb.Append(c);
                else
                    sb.Append('_');
            }

            string result = sb.ToString().Trim('_');
            if (string.IsNullOrEmpty(result)) result = "col";
            if (char.IsDigit(result[0])) result = $"c_{result}";
            return result;
        }
    }
}
