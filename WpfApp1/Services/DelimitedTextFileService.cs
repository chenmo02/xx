using System.Data;
using System.Globalization;
using System.IO;
using System.Text;
using CsvHelper;
using CsvHelper.Configuration;

namespace WpfApp1.Services
{
    public sealed class DelimitedTextFileInfo
    {
        public required string FilePath { get; init; }

        public required string FileName { get; init; }

        public required long FileSize { get; init; }

        public required Encoding Encoding { get; init; }

        public required string Delimiter { get; init; }

        public required int TotalRows { get; init; }

        public required int TotalColumns { get; init; }

        public required IReadOnlyList<string> Columns { get; init; }
    }

    public sealed class DelimitedTextLoadResult
    {
        public required string FilePath { get; init; }

        public required string FileName { get; init; }

        public required long FileSize { get; init; }

        public required Encoding Encoding { get; init; }

        public required string Delimiter { get; init; }

        public required DataTable Table { get; init; }
    }

    public static class DelimitedTextFileService
    {
        static DelimitedTextFileService()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }

        public static DelimitedTextFileInfo GetFileInfo(string filePath)
        {
            string normalizedPath = Path.GetFullPath(filePath);
            var fileInfo = new FileInfo(normalizedPath);
            Encoding encoding = DetectEncoding(normalizedPath);
            string delimiter = DetectDelimiter(normalizedPath, encoding);
            var columns = new List<string>();
            int totalRows = 0;

            using var stream = new FileStream(normalizedPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var reader = new StreamReader(stream, encoding, detectEncodingFromByteOrderMarks: true);
            using var csv = new CsvReader(reader, CreateConfiguration(delimiter));

            if (csv.Read())
            {
                csv.ReadHeader();
                columns = BuildNormalizedHeaders(csv.HeaderRecord).ToList();

                while (csv.Read())
                {
                    totalRows++;
                }
            }

            return new DelimitedTextFileInfo
            {
                FilePath = normalizedPath,
                FileName = fileInfo.Name,
                FileSize = fileInfo.Length,
                Encoding = encoding,
                Delimiter = delimiter,
                TotalRows = totalRows,
                TotalColumns = columns.Count,
                Columns = columns
            };
        }

        public static DelimitedTextLoadResult LoadFile(string filePath, int previewRows = 0)
        {
            string normalizedPath = Path.GetFullPath(filePath);
            var fileInfo = new FileInfo(normalizedPath);
            Encoding encoding = DetectEncoding(normalizedPath);
            string delimiter = DetectDelimiter(normalizedPath, encoding);
            DataTable table = ParseDelimitedFile(normalizedPath, encoding, delimiter, previewRows);

            return new DelimitedTextLoadResult
            {
                FilePath = normalizedPath,
                FileName = fileInfo.Name,
                FileSize = fileInfo.Length,
                Encoding = encoding,
                Delimiter = delimiter,
                Table = table
            };
        }

        public static string GetDelimiterName(string delimiter) => delimiter switch
        {
            "," => "逗号",
            "\t" => "Tab",
            ";" => "分号",
            "|" => "竖线",
            _ => delimiter
        };

        private static DataTable ParseDelimitedFile(string filePath, Encoding encoding, string delimiter, int previewRows)
        {
            var table = new DataTable(Path.GetFileNameWithoutExtension(filePath));

            using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var reader = new StreamReader(stream, encoding, detectEncodingFromByteOrderMarks: true);
            using var csv = new CsvReader(reader, CreateConfiguration(delimiter));

            if (!csv.Read())
            {
                return table;
            }

            csv.ReadHeader();
            IReadOnlyList<string> headers = BuildNormalizedHeaders(csv.HeaderRecord);
            for (int columnIndex = 0; columnIndex < headers.Count; columnIndex++)
            {
                table.Columns.Add(headers[columnIndex], typeof(string));
            }

            int rowCount = 0;
            while (csv.Read())
            {
                if (previewRows > 0 && rowCount >= previewRows)
                {
                    break;
                }

                DataRow row = table.NewRow();
                for (int columnIndex = 0; columnIndex < table.Columns.Count; columnIndex++)
                {
                    try
                    {
                        row[columnIndex] = csv.GetField(columnIndex) ?? string.Empty;
                    }
                    catch
                    {
                        row[columnIndex] = string.Empty;
                    }
                }

                table.Rows.Add(row);
                rowCount++;
            }

            return table;
        }

        private static CsvConfiguration CreateConfiguration(string delimiter)
        {
            return new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                Delimiter = delimiter,
                HasHeaderRecord = true,
                MissingFieldFound = null,
                BadDataFound = null,
                TrimOptions = TrimOptions.Trim
            };
        }

        private static IReadOnlyList<string> BuildNormalizedHeaders(string[]? headers)
        {
            headers ??= [];
            var normalizedHeaders = new List<string>(headers.Length);
            var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            for (int index = 0; index < headers.Length; index++)
            {
                string baseName = string.IsNullOrWhiteSpace(headers[index]) ? $"Column{index + 1}" : headers[index].Trim();
                string columnName = baseName;
                int suffix = 2;

                while (!usedNames.Add(columnName))
                {
                    columnName = $"{baseName}_{suffix}";
                    suffix++;
                }

                normalizedHeaders.Add(columnName);
            }

            return normalizedHeaders;
        }

        private static Encoding DetectEncoding(string filePath)
        {
            var bom = new byte[4];
            using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                _ = stream.ReadAtLeast(bom, 4, throwOnEndOfStream: false);
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
                using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                using var reader = new StreamReader(stream, utf8, detectEncodingFromByteOrderMarks: true);
                _ = reader.ReadToEnd();
                return Encoding.UTF8;
            }
            catch
            {
                return Encoding.GetEncoding("GBK");
            }
        }

        private static string DetectDelimiter(string filePath, Encoding encoding)
        {
            using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var reader = new StreamReader(stream, encoding, detectEncodingFromByteOrderMarks: true);
            string? firstLine = reader.ReadLine();

            if (string.IsNullOrEmpty(firstLine))
            {
                return ",";
            }

            int commaCount = firstLine.Count(character => character == ',');
            int tabCount = firstLine.Count(character => character == '\t');
            int semicolonCount = firstLine.Count(character => character == ';');
            int pipeCount = firstLine.Count(character => character == '|');
            int maxCount = Math.Max(Math.Max(commaCount, tabCount), Math.Max(semicolonCount, pipeCount));

            if (maxCount == 0)
            {
                return ",";
            }

            if (maxCount == tabCount)
            {
                return "\t";
            }

            if (maxCount == semicolonCount)
            {
                return ";";
            }

            if (maxCount == pipeCount)
            {
                return "|";
            }

            return ",";
        }
    }
}
