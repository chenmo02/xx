using System.Data;
using System.IO;
using System.Text;

namespace WpfApp1.Services
{
    public static class ExportService
    {
        public static void ExportSql(string filePath, string sql)
        {
            File.WriteAllText(filePath, sql, new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));
        }

        public static void ExportCsv(string filePath, DataTable table)
        {
            var sb = new StringBuilder();
            sb.AppendLine(string.Join(",", table.Columns.Cast<DataColumn>().Select(column => QuoteCsv(column.ColumnName))));

            foreach (DataRow row in table.Rows)
            {
                sb.AppendLine(string.Join(",", row.ItemArray.Select(value => QuoteCsv(value == DBNull.Value ? "" : value?.ToString() ?? ""))));
            }

            File.WriteAllText(filePath, sb.ToString(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));
        }

        public static void ExportJson(string filePath, DataTable table)
        {
            var sb = new StringBuilder();
            sb.AppendLine("[");

            for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
            {
                DataRow row = table.Rows[rowIndex];
                sb.AppendLine("  {");

                for (int colIndex = 0; colIndex < table.Columns.Count; colIndex++)
                {
                    DataColumn column = table.Columns[colIndex];
                    object value = row[column];
                    string jsonValue = FormatJsonValue(value);

                    sb.Append($"    \"{EscapeJson(column.ColumnName)}\": {jsonValue}");
                    sb.AppendLine(colIndex < table.Columns.Count - 1 ? "," : string.Empty);
                }

                sb.Append("  }");
                sb.AppendLine(rowIndex < table.Rows.Count - 1 ? "," : string.Empty);
            }

            sb.AppendLine("]");
            File.WriteAllText(filePath, sb.ToString(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));
        }

        private static string QuoteCsv(string value) => $"\"{value.Replace("\"", "\"\"")}\"";

        private static string FormatJsonValue(object value)
        {
            if (value == DBNull.Value || value == null)
            {
                return "null";
            }

            return value switch
            {
                bool boolValue => boolValue ? "true" : "false",
                byte or short or int or long or float or double or decimal => Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture) ?? "null",
                DateTime dateTime => $"\"{dateTime:yyyy-MM-ddTHH:mm:ss}\"",
                _ => $"\"{EscapeJson(value.ToString() ?? "")}\""
            };
        }

        private static string EscapeJson(string value)
        {
            return value
                .Replace("\\", "\\\\")
                .Replace("\"", "\\\"")
                .Replace("\r", "\\r")
                .Replace("\n", "\\n")
                .Replace("\t", "\\t");
        }
    }
}
