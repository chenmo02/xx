using System.Data;
using System.Globalization;
using System.IO;
using System.Text;
using CsvHelper;
using CsvHelper.Configuration;
using ExcelDataReader;

namespace WpfApp1.Services
{
    public sealed class ImportProgressInfo
    {
        public string Stage { get; init; } = "";

        public int Percentage { get; init; }

        public int Current { get; init; }

        public int Total { get; init; }
    }

    public static class FileParserService
    {
        private sealed class DbfField
        {
            public required string Name { get; init; }

            public required char Type { get; init; }

            public required byte Length { get; init; }

            public required byte DecimalCount { get; init; }
        }

        private static readonly IReadOnlyList<string> DbfEncodings =
        [
            "UTF-8",
            "GB2312",
            "GBK",
            "GB18030",
            "Big5",
            "UTF-16",
            "ASCII",
            "Windows-1252"
        ];

        static FileParserService()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }

        public static IReadOnlyList<string> GetDbfEncodings() => DbfEncodings;

        public static List<string> GetSheetNames(string filePath)
        {
            string ext = Path.GetExtension(filePath).ToLowerInvariant();
            if (ext is not (".xlsx" or ".xls"))
            {
                return [];
            }

            using var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var reader = ExcelReaderFactory.CreateReader(stream);
            var dataSet = reader.AsDataSet();
            return dataSet.Tables.Cast<DataTable>().Select(table => table.TableName).ToList();
        }

        public static (long fileSize, int totalRows, int totalCols) GetFileInfo(
            string filePath,
            string? sheetName = null,
            string? dbfEncoding = null)
        {
            var fileInfo = new FileInfo(filePath);
            long fileSize = fileInfo.Length;
            string ext = Path.GetExtension(filePath).ToLowerInvariant();

            return ext switch
            {
                ".xlsx" or ".xls" => GetExcelFileInfo(filePath, sheetName, fileSize),
                ".csv" => GetCsvFileInfo(filePath, fileSize),
                ".dbf" => GetDbfFileInfo(filePath, fileSize, dbfEncoding),
                _ => throw new NotSupportedException($"不支持的文件格式: {ext}")
            };
        }

        public static DataTable ParseFile(
            string filePath,
            string? sheetName = null,
            int previewRows = 0,
            string? dbfEncoding = null,
            IProgress<ImportProgressInfo>? progress = null)
        {
            string ext = Path.GetExtension(filePath).ToLowerInvariant();

            return ext switch
            {
                ".xlsx" or ".xls" => ParseExcel(filePath, sheetName, previewRows, progress),
                ".csv" => ParseCsv(filePath, previewRows),
                ".dbf" => ParseDbf(filePath, previewRows, dbfEncoding, progress),
                _ => throw new NotSupportedException($"不支持的文件格式: {ext}")
            };
        }

        private static (long fileSize, int totalRows, int totalCols) GetExcelFileInfo(string filePath, string? sheetName, long fileSize)
        {
            DataTable table = ParseExcel(filePath, sheetName, previewRows: 0, progress: null);
            return (fileSize, table.Rows.Count, table.Columns.Count);
        }

        private static (long fileSize, int totalRows, int totalCols) GetCsvFileInfo(string filePath, long fileSize)
        {
            Encoding encoding = DetectEncoding(filePath);
            int lineCount = 0;

            using (var sr = new StreamReader(filePath, encoding))
            {
                while (sr.ReadLine() != null)
                {
                    lineCount++;
                }
            }

            DataTable headerTable = ParseCsv(filePath, previewRows: 1);
            return (fileSize, Math.Max(0, lineCount - 1), headerTable.Columns.Count);
        }

        private static (long fileSize, int totalRows, int totalCols) GetDbfFileInfo(string filePath, long fileSize, string? dbfEncoding)
        {
            using var input = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var reader = new BinaryReader(input);

            _ = reader.ReadByte();
            _ = reader.ReadBytes(3);
            int totalRows = reader.ReadInt32();
            short headerLength = reader.ReadInt16();
            _ = reader.ReadInt16();
            _ = reader.ReadBytes(20);

            int columnCount = (headerLength - 33) / 32;
            return (fileSize, totalRows, Math.Max(columnCount, 0));
        }

        private static DataTable ParseExcel(
            string filePath,
            string? sheetName,
            int previewRows,
            IProgress<ImportProgressInfo>? progress)
        {
            progress?.Report(new ImportProgressInfo
            {
                Stage = "正在读取 Excel 工作簿...",
                Percentage = 5
            });

            using var stream = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var reader = ExcelReaderFactory.CreateReader(stream);
            string? expectedSheetName = sheetName;
            var dataSet = reader.AsDataSet(new ExcelDataSetConfiguration
            {
                FilterSheet = (tableReader, sheetIndex) =>
                {
                    if (!string.IsNullOrWhiteSpace(expectedSheetName))
                    {
                        return string.Equals(tableReader.Name, expectedSheetName, StringComparison.Ordinal);
                    }

                    return sheetIndex == 0;
                },
                ConfigureDataTable = _ => new ExcelDataTableConfiguration
                {
                    UseHeaderRow = true
                }
            });

            string? selectedSheet = sheetName;
            if (string.IsNullOrWhiteSpace(selectedSheet))
            {
                selectedSheet = dataSet.Tables.Count > 0 ? dataSet.Tables[0].TableName : null;
            }

            if (string.IsNullOrWhiteSpace(selectedSheet))
            {
                throw new InvalidOperationException("Excel 文件中没有可用工作表。");
            }

            var sourceTable = dataSet.Tables[selectedSheet];
            if (sourceTable == null)
            {
                throw new InvalidOperationException($"找不到工作表: {selectedSheet}");
            }

            progress?.Report(new ImportProgressInfo
            {
                Stage = $"正在加载工作表 {selectedSheet}...",
                Percentage = 50
            });

            var normalized = NormalizeTable(sourceTable, previewRows);

            progress?.Report(new ImportProgressInfo
            {
                Stage = "Excel 读取完成",
                Percentage = 100,
                Current = normalized.Rows.Count,
                Total = normalized.Rows.Count
            });

            return normalized;
        }

        private static DataTable ParseCsv(string filePath, int previewRows)
        {
            var dt = new DataTable();
            Encoding encoding = DetectEncoding(filePath);
            string delimiter = DetectDelimiter(filePath, encoding);

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

            if (!csv.Read())
            {
                return dt;
            }

            csv.ReadHeader();
            string[] headers = csv.HeaderRecord ?? [];

            for (int i = 0; i < headers.Length; i++)
            {
                dt.Columns.Add(GetUniqueColumnName(dt, headers[i], i + 1), typeof(string));
            }

            if (dt.Columns.Count == 0)
            {
                return dt;
            }

            int rowCount = 0;
            while (csv.Read())
            {
                if (previewRows > 0 && rowCount >= previewRows)
                {
                    break;
                }

                var row = dt.NewRow();
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    try
                    {
                        row[i] = csv.GetField(i) ?? "";
                    }
                    catch
                    {
                        row[i] = "";
                    }
                }

                dt.Rows.Add(row);
                rowCount++;
            }

            return dt;
        }

        private static DataTable ParseDbf(
            string filePath,
            int previewRows,
            string? dbfEncoding,
            IProgress<ImportProgressInfo>? progress)
        {
            string encodingName = string.IsNullOrWhiteSpace(dbfEncoding) ? "UTF-8" : dbfEncoding;
            Encoding encoding = ResolveDbfEncoding(encodingName);

            var table = new DataTable(Path.GetFileNameWithoutExtension(filePath));

            using var input = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var reader = new BinaryReader(input);

            _ = reader.ReadByte();
            _ = reader.ReadBytes(3);
            int totalRows = reader.ReadInt32();
            short headerLength = reader.ReadInt16();
            short recordLength = reader.ReadInt16();
            _ = reader.ReadBytes(20);

            int fieldCount = (headerLength - 33) / 32;
            var fields = new List<DbfField>(fieldCount);

            for (int i = 0; i < fieldCount; i++)
            {
                byte[] nameBytes = reader.ReadBytes(11);
                string rawName = Encoding.ASCII.GetString(nameBytes).TrimEnd('\0', ' ');
                char fieldType = (char)reader.ReadByte();
                _ = reader.ReadBytes(4);
                byte length = reader.ReadByte();
                byte decimalCount = reader.ReadByte();
                _ = reader.ReadBytes(14);

                string columnName = GetUniqueColumnName(table, rawName, i + 1);
                table.Columns.Add(columnName, GetDbfColumnType(fieldType));
                fields.Add(new DbfField
                {
                    Name = columnName,
                    Type = fieldType,
                    Length = length,
                    DecimalCount = decimalCount
                });
            }

            _ = reader.ReadByte();

            int maxRows = previewRows > 0 ? Math.Min(previewRows, totalRows) : totalRows;
            for (int rowIndex = 0; rowIndex < totalRows; rowIndex++)
            {
                byte deletedFlag = reader.ReadByte();
                if (deletedFlag == 0x2A)
                {
                    _ = reader.ReadBytes(recordLength - 1);
                    continue;
                }

                if (table.Rows.Count >= maxRows)
                {
                    _ = reader.ReadBytes(recordLength - 1);
                    continue;
                }

                var row = table.NewRow();
                foreach (DbfField field in fields)
                {
                    byte[] rawValue = reader.ReadBytes(field.Length);
                    string value = encoding.GetString(rawValue).Trim();
                    row[field.Name] = ConvertDbfValue(value, field.Type, field.DecimalCount);
                }

                table.Rows.Add(row);

                if (progress != null && totalRows > 0 && (rowIndex + 1 == totalRows || (rowIndex + 1) % 2000 == 0))
                {
                    progress.Report(new ImportProgressInfo
                    {
                        Stage = "正在读取 DBF 文件...",
                        Current = rowIndex + 1,
                        Total = totalRows,
                        Percentage = Math.Min(100, (rowIndex + 1) * 100 / totalRows)
                    });
                }
            }

            progress?.Report(new ImportProgressInfo
            {
                Stage = $"DBF 读取完成 (编码: {encoding.WebName})",
                Current = table.Rows.Count,
                Total = maxRows,
                Percentage = 100
            });

            return table;
        }

        private static DataTable NormalizeTable(DataTable source, int previewRows)
        {
            var normalized = new DataTable(source.TableName);

            for (int i = 0; i < source.Columns.Count; i++)
            {
                DataColumn sourceColumn = source.Columns[i];
                string columnName = GetUniqueColumnName(normalized, sourceColumn.ColumnName, i + 1);
                Type columnType = sourceColumn.DataType == typeof(DBNull) ? typeof(string) : sourceColumn.DataType;
                normalized.Columns.Add(columnName, columnType);
            }

            int maxRows = previewRows > 0 ? Math.Min(previewRows, source.Rows.Count) : source.Rows.Count;
            for (int rowIndex = 0; rowIndex < maxRows; rowIndex++)
            {
                DataRow newRow = normalized.NewRow();
                for (int colIndex = 0; colIndex < normalized.Columns.Count; colIndex++)
                {
                    object value = source.Rows[rowIndex][colIndex];
                    newRow[colIndex] = value == DBNull.Value ? DBNull.Value : value;
                }

                normalized.Rows.Add(newRow);
            }

            return normalized;
        }

        private static string GetUniqueColumnName(DataTable table, string? rawName, int index)
        {
            string baseName = string.IsNullOrWhiteSpace(rawName) ? $"Column{index}" : rawName.Trim();
            string columnName = baseName;
            int suffix = 1;

            while (table.Columns.Contains(columnName))
            {
                suffix++;
                columnName = $"{baseName}_{suffix}";
            }

            return columnName;
        }

        private static Encoding DetectEncoding(string filePath)
        {
            var bom = new byte[4];
            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                _ = fs.ReadAtLeast(bom, 4, throwOnEndOfStream: false);
            }

            if (bom[0] == 0xEF && bom[1] == 0xBB && bom[2] == 0xBF)
            {
                return Encoding.UTF8;
            }

            if (bom[0] == 0xFF && bom[1] == 0xFE)
            {
                return Encoding.Unicode;
            }

            if (bom[0] == 0xFE && bom[1] == 0xFF)
            {
                return Encoding.BigEndianUnicode;
            }

            try
            {
                var utf8 = new UTF8Encoding(false, true);
                using var sr = new StreamReader(filePath, utf8, true);
                _ = sr.ReadToEnd();
                return Encoding.UTF8;
            }
            catch
            {
                return Encoding.GetEncoding("GBK");
            }
        }

        private static string DetectDelimiter(string filePath, Encoding encoding)
        {
            using var reader = new StreamReader(filePath, encoding);
            string? firstLine = reader.ReadLine();

            if (string.IsNullOrEmpty(firstLine))
            {
                return ",";
            }

            int commaCount = firstLine.Count(c => c == ',');
            int tabCount = firstLine.Count(c => c == '\t');
            int semicolonCount = firstLine.Count(c => c == ';');
            int pipeCount = firstLine.Count(c => c == '|');

            int max = Math.Max(Math.Max(commaCount, tabCount), Math.Max(semicolonCount, pipeCount));
            if (max == 0)
            {
                return ",";
            }

            if (max == tabCount)
            {
                return "\t";
            }

            if (max == semicolonCount)
            {
                return ";";
            }

            if (max == pipeCount)
            {
                return "|";
            }

            return ",";
        }

        private static Encoding ResolveDbfEncoding(string encodingName)
        {
            try
            {
                return Encoding.GetEncoding(encodingName);
            }
            catch
            {
                return Encoding.UTF8;
            }
        }

        private static Type GetDbfColumnType(char dbfType) => dbfType switch
        {
            'C' => typeof(string),
            'M' => typeof(string),
            'F' => typeof(decimal),
            'N' => typeof(decimal),
            'L' => typeof(bool),
            'D' => typeof(DateTime),
            'I' => typeof(int),
            _ => typeof(string)
        };

        private static object ConvertDbfValue(string value, char fieldType, byte decimalCount)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return DBNull.Value;
            }

            try
            {
                return fieldType switch
                {
                    'C' or 'M' => value,
                    'F' or 'N' => decimal.TryParse(value, out decimal decimalValue) ? decimalValue : DBNull.Value,
                    'L' => ConvertDbfBoolean(value),
                    'D' => ConvertDbfDate(value),
                    'I' => int.TryParse(value, out int intValue) ? intValue : DBNull.Value,
                    _ => value
                };
            }
            catch
            {
                return DBNull.Value;
            }
        }

        private static object ConvertDbfBoolean(string value)
        {
            string upper = value.Trim().ToUpperInvariant();
            if (upper is "T" or "Y" or "1")
            {
                return true;
            }

            if (upper is "F" or "N" or "0")
            {
                return false;
            }

            return DBNull.Value;
        }

        private static object ConvertDbfDate(string value)
        {
            if (value.Length == 8 &&
                int.TryParse(value[..4], out int year) &&
                int.TryParse(value.Substring(4, 2), out int month) &&
                int.TryParse(value.Substring(6, 2), out int day))
            {
                return new DateTime(year, month, day);
            }

            return DBNull.Value;
        }
    }
}
