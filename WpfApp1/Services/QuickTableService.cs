using System.Globalization;
using System.Text;

namespace WpfApp1.Services;

public sealed class QuickTableColumn
{
    public int RowNumber { get; set; }
    public string OriginalName { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
    private string _type = "varchar";
    public string Type
    {
        get => _type;
        set
        {
            _type = QuickTableService.NormalizeType(value);
            if (_type == "varchar")
                Length = 1000;
        }
    }
    public int Length { get; set; } = 255;
    public int Precision { get; set; } = 18;
    public int Scale { get; set; } = 6;
    public bool Nullable { get; set; } = true;
    public string Example { get; set; } = string.Empty;
    public string Inference { get; set; } = string.Empty;
}

public static class QuickTableService
{
    public static readonly string[] TypeChoices = ["varchar", "text", "bigint", "numeric", "timestamp", "boolean"];

    public static string NormalizeType(string? type)
    {
        string value = type?.Trim().ToLowerInvariant() ?? string.Empty;
        return value switch
        {
            "varchar" or "varchar2" or "nvarchar" or "varcahr" => "varchar",
            "text" or "longtext" or "clob" => "text",
            "bigint" or "int" or "integer" or "number" => "bigint",
            "numeric" or "decimal" => "numeric",
            "timestamp" or "datetime" or "datetime2" or "date" => "timestamp",
            "boolean" or "bool" or "bit" => "boolean",
            _ => "varchar"
        };
    }

    public static List<QuickTableColumn> InferColumns(IReadOnlyList<string> headers, IReadOnlyList<IReadOnlyList<string?>> rows)
    {
        var result = new List<QuickTableColumn>(headers.Count);
        for (int i = 0; i < headers.Count; i++)
        {
            var values = rows.Select(r => i < r.Count ? r[i] : null).Where(v => !string.IsNullOrWhiteSpace(v)).ToList();
            string type = NormalizeType(InferType(values));
            int maxLength = type == "varchar"
                ? 1000
                : Math.Max(1, values.Count == 0 ? 255 : values.Max(v => v!.Length));
            result.Add(new QuickTableColumn
            {
                RowNumber = i + 1,
                OriginalName = headers[i],
                Name = headers[i],
                Type = type,
                Length = Math.Min(Math.Max(maxLength, 1), 4000),
                Example = values.FirstOrDefault() ?? string.Empty,
                Inference = values.Count == 0 ? "无非空样本，默认 varchar" : $"根据 {values.Count} 个非空值推断"
            });
        }
        return result;
    }

    public static string GenerateSql(SqlGeneratorService.DbType dbType, string tableName,
        IReadOnlyList<QuickTableColumn> columns, bool dropIfExists, bool quoteIdentifiers)
    {
        if (string.IsNullOrWhiteSpace(tableName)) throw new InvalidOperationException("表名不能为空。");
        if (columns.Count == 0) throw new InvalidOperationException("至少需要一个字段。");
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var column in columns)
        {
            if (string.IsNullOrWhiteSpace(column.Name)) throw new InvalidOperationException("字段名不能为空。");
            if (!names.Add(column.Name.Trim())) throw new InvalidOperationException($"字段名大小写转换后发生重名：{column.Name}");
        }

        string wrappedTable = quoteIdentifiers ? QuoteCompositeName(dbType, tableName.Trim()) : tableName.Trim();
        var sb = new StringBuilder();
        if (dropIfExists)
        {
            if (dbType is SqlGeneratorService.DbType.PostgreSQL or SqlGeneratorService.DbType.MySQL)
                sb.AppendLine($"DROP TABLE IF EXISTS {wrappedTable};");
            else if (dbType == SqlGeneratorService.DbType.SqlServer)
                sb.AppendLine($"IF OBJECT_ID(N'{tableName.Replace("'", "''")}', 'U') IS NOT NULL DROP TABLE {wrappedTable};");
            else
            {
                sb.AppendLine("BEGIN");
                sb.AppendLine($"    EXECUTE IMMEDIATE 'DROP TABLE {tableName.Replace("'", "''")}';");
                sb.AppendLine("EXCEPTION WHEN OTHERS THEN IF SQLCODE != -942 THEN RAISE; END IF;");
                sb.AppendLine("END;");
                sb.AppendLine("/");
            }
            sb.AppendLine();
        }

        sb.AppendLine($"CREATE TABLE {wrappedTable} (");
        for (int i = 0; i < columns.Count; i++)
        {
            var column = columns[i];
            string name = quoteIdentifiers ? QuoteName(dbType, column.Name) : column.Name;
            sb.Append($"    {name} {MapType(dbType, column)} {(column.Nullable ? "NULL" : "NOT NULL")}");
            sb.AppendLine(i == columns.Count - 1 ? string.Empty : ",");
        }
        sb.AppendLine(");");
        return sb.ToString();
    }

    public static string ApplyCase(string name, string mode) => mode switch
    {
        "大写" => name.ToUpperInvariant(),
        "小写" => name.ToLowerInvariant(),
        _ => name
    };

    private static string InferType(List<string?> values)
    {
        if (values.Count == 0) return "varchar";
        if (values.All(v => bool.TryParse(v, out _) || v!.Equals("0") || v.Equals("1"))) return "boolean";
        if (values.All(v => long.TryParse(v, NumberStyles.Integer, CultureInfo.InvariantCulture, out _))) return "bigint";
        if (values.All(v => decimal.TryParse(v, NumberStyles.Number, CultureInfo.InvariantCulture, out _))) return "numeric";
        if (values.All(v => DateTime.TryParse(v, CultureInfo.CurrentCulture, DateTimeStyles.None, out _))) return "timestamp";
        return values.Any(v => v!.Length > 255) ? "text" : "varchar";
    }

    private static string MapType(SqlGeneratorService.DbType dbType, QuickTableColumn c) => NormalizeType(c.Type) switch
    {
        "text" => dbType switch
        {
            SqlGeneratorService.DbType.SqlServer => "NVARCHAR(MAX)",
            SqlGeneratorService.DbType.MySQL => "LONGTEXT",
            SqlGeneratorService.DbType.Oracle => "CLOB",
            _ => "TEXT"
        },
        "varchar" => dbType switch
        {
            SqlGeneratorService.DbType.SqlServer => $"NVARCHAR({Math.Max(1, c.Length)})",
            SqlGeneratorService.DbType.Oracle => $"VARCHAR2({Math.Max(1, c.Length)})",
            _ => $"VARCHAR({Math.Max(1, c.Length)})"
        },
        "bigint" => dbType == SqlGeneratorService.DbType.Oracle ? "NUMBER(19)" : "BIGINT",
        "numeric" => dbType == SqlGeneratorService.DbType.Oracle ? $"NUMBER({c.Precision},{c.Scale})" : $"NUMERIC({c.Precision},{c.Scale})",
        "timestamp" => dbType switch
        {
            SqlGeneratorService.DbType.SqlServer => "DATETIME2",
            SqlGeneratorService.DbType.MySQL => "DATETIME",
            SqlGeneratorService.DbType.Oracle => "DATE",
            _ => "TIMESTAMP"
        },
        "boolean" => dbType switch
        {
            SqlGeneratorService.DbType.SqlServer => "BIT",
            SqlGeneratorService.DbType.MySQL => "TINYINT(1)",
            SqlGeneratorService.DbType.Oracle => "NUMBER(1)",
            _ => "BOOLEAN"
        },
        _ => "VARCHAR(255)"
    };

    private static string QuoteCompositeName(SqlGeneratorService.DbType dbType, string name) =>
        string.Join('.', name.Split('.', StringSplitOptions.RemoveEmptyEntries).Select(part => QuoteName(dbType, part)));

    private static string QuoteName(SqlGeneratorService.DbType dbType, string name)
    {
        string clean = name.Trim().Trim('[', ']', '`', '"');
        return dbType switch
        {
            SqlGeneratorService.DbType.MySQL => $"`{clean.Replace("`", "``")}`",
            SqlGeneratorService.DbType.SqlServer => $"[{clean.Replace("]", "]]" )}]",
            _ => $"\"{clean.Replace("\"", "\"\"")}\""
        };
    }
}
