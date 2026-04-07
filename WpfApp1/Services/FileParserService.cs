using System.Data;
using System.Globalization;
using System.IO;
using System.Text;
using CsvHelper;
using CsvHelper.Configuration;
using OfficeOpenXml;

namespace WpfApp1.Services
{
    /// <summary>
    /// 文件解析服务：支持 Excel (.xlsx) 和 CSV 文件解析为 DataTable
    /// </summary>
    public static class FileParserService
    {
        static FileParserService()
        {
            ExcelPackage.License.SetNonCommercialOrganization("DM Tools");
        }

        /// <summary>
        /// 获取 Excel 文件中的所有 Sheet 名称
        /// </summary>
        public static List<string> GetSheetNames(string filePath)
        {
            string ext = Path.GetExtension(filePath).ToLower();
            if (ext is not (".xlsx" or ".xls"))
                return ["Sheet1 (CSV)"];

            using var package = new ExcelPackage(new FileInfo(filePath));
            return package.Workbook.Worksheets.Select(ws => ws.Name).ToList();
        }

        /// <summary>
        /// 解析指定 Sheet 的数据
        /// </summary>
        public static DataTable ParseFile(string filePath, string? sheetName = null, int previewRows = 0)
        {
            string ext = Path.GetExtension(filePath).ToLower();
            return ext switch
            {
                ".xlsx" or ".xls" => ParseExcel(filePath, sheetName, previewRows),
                ".csv" => ParseCsv(filePath, previewRows),
                _ => throw new NotSupportedException($"不支持的文件格式: {ext}")
            };
        }

        /// <summary>
        /// 获取指定 Sheet 的行列信息
        /// </summary>
        public static (long fileSize, int totalRows, int totalCols) GetFileInfo(string filePath, string? sheetName = null)
        {
            var fileInfo = new FileInfo(filePath);
            long fileSize = fileInfo.Length;

            string ext = Path.GetExtension(filePath).ToLower();
            if (ext is ".xlsx" or ".xls")
            {
                using var package = new ExcelPackage(new FileInfo(filePath));
                var ws = string.IsNullOrEmpty(sheetName)
                    ? package.Workbook.Worksheets[0]
                    : package.Workbook.Worksheets[sheetName];

                if (ws?.Dimension == null) return (fileSize, 0, 0);
                return (fileSize, ws.Dimension.Rows - 1, ws.Dimension.Columns);
            }
            else
            {
                Encoding encoding = DetectEncoding(filePath);
                int lineCount = 0;
                using (var sr = new StreamReader(filePath, encoding))
                {
                    while (sr.ReadLine() != null) lineCount++;
                }
                var dt = ParseCsv(filePath, 1);
                return (fileSize, Math.Max(0, lineCount - 1), dt.Columns.Count);
            }
        }

        /// <summary>
        /// 解析 Excel 文件指定 Sheet
        /// </summary>
        private static DataTable ParseExcel(string filePath, string? sheetName, int previewRows)
        {
            var dt = new DataTable();

            using var package = new ExcelPackage(new FileInfo(filePath));

            var worksheet = string.IsNullOrEmpty(sheetName)
                ? package.Workbook.Worksheets[0]
                : package.Workbook.Worksheets[sheetName];

            if (worksheet == null)
                throw new Exception($"找不到工作表: {sheetName}");

            if (worksheet.Dimension == null)
                throw new Exception($"工作表 [{worksheet.Name}] 为空，没有数据");

            int totalRows = worksheet.Dimension.Rows;
            int totalCols = worksheet.Dimension.Columns;

            // 第一行作为列名
            for (int col = 1; col <= totalCols; col++)
            {
                string colName = worksheet.Cells[1, col].Text?.Trim() ?? $"Column{col}";
                if (string.IsNullOrWhiteSpace(colName))
                    colName = $"Column{col}";
                if (dt.Columns.Contains(colName))
                    colName = $"{colName}_{col}";
                dt.Columns.Add(colName, typeof(string));
            }

            // 读取数据行
            int maxRows = previewRows > 0 ? Math.Min(previewRows + 1, totalRows) : totalRows;
            for (int row = 2; row <= maxRows; row++)
            {
                var dataRow = dt.NewRow();
                for (int col = 1; col <= totalCols; col++)
                {
                    dataRow[col - 1] = worksheet.Cells[row, col].Text ?? "";
                }
                dt.Rows.Add(dataRow);
            }

            return dt;
        }

        /// <summary>
        /// 解析 CSV 文件（自动检测编码和分隔符）
        /// </summary>
        private static DataTable ParseCsv(string filePath, int previewRows)
        {
            var dt = new DataTable();
            var encoding = DetectEncoding(filePath);
            var delimiter = DetectDelimiter(filePath, encoding);

            var config = new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                Delimiter = delimiter,
                HasHeaderRecord = true,
                MissingFieldFound = null,
                BadDataFound = null,
                TrimOptions = TrimOptions.Trim
            };

            using var reader = new StreamReader(filePath, encoding);
            using var csv = new CsvReader(reader, config);

            csv.Read();
            csv.ReadHeader();

            if (csv.HeaderRecord != null)
            {
                for (int i = 0; i < csv.HeaderRecord.Length; i++)
                {
                    string colName = csv.HeaderRecord[i]?.Trim() ?? $"Column{i + 1}";
                    if (string.IsNullOrWhiteSpace(colName))
                        colName = $"Column{i + 1}";
                    if (dt.Columns.Contains(colName))
                        colName = $"{colName}_{i + 1}";
                    dt.Columns.Add(colName, typeof(string));
                }
            }

            int rowCount = 0;
            while (csv.Read())
            {
                if (previewRows > 0 && rowCount >= previewRows)
                    break;

                var dataRow = dt.NewRow();
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    try { dataRow[i] = csv.GetField(i) ?? ""; }
                    catch { dataRow[i] = ""; }
                }
                dt.Rows.Add(dataRow);
                rowCount++;
            }

            return dt;
        }

        /// <summary>
        /// 检测文件编码
        /// </summary>
        private static Encoding DetectEncoding(string filePath)
        {
            var bom = new byte[4];
            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                _ = fs.ReadAtLeast(bom, 4, throwOnEndOfStream: false);
            }

            if (bom[0] == 0xEF && bom[1] == 0xBB && bom[2] == 0xBF) return Encoding.UTF8;
            if (bom[0] == 0xFF && bom[1] == 0xFE) return Encoding.Unicode;
            if (bom[0] == 0xFE && bom[1] == 0xFF) return Encoding.BigEndianUnicode;

            try
            {
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                var utf8 = new UTF8Encoding(false, true);
                using var sr = new StreamReader(filePath, utf8, true);
                sr.ReadToEnd();
                return Encoding.UTF8;
            }
            catch
            {
                return Encoding.GetEncoding("GBK");
            }
        }

        /// <summary>
        /// 自动检测 CSV 分隔符
        /// </summary>
        private static string DetectDelimiter(string filePath, Encoding encoding)
        {
            using var reader = new StreamReader(filePath, encoding);
            string? firstLine = reader.ReadLine();

            if (string.IsNullOrEmpty(firstLine)) return ",";

            int commaCount = firstLine.Count(c => c == ',');
            int tabCount = firstLine.Count(c => c == '\t');
            int semicolonCount = firstLine.Count(c => c == ';');
            int pipeCount = firstLine.Count(c => c == '|');

            int max = Math.Max(Math.Max(commaCount, tabCount), Math.Max(semicolonCount, pipeCount));

            if (max == 0) return ",";
            if (max == tabCount) return "\t";
            if (max == semicolonCount) return ";";
            if (max == pipeCount) return "|";
            return ",";
        }
    }
}
