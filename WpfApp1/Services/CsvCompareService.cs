using System.Data;
using System.Globalization;
using System.Text;

namespace WpfApp1.Services
{
    public enum CsvCompareMode
    {
        ByRowNumber,
        ByKeyColumns
    }

    public enum CsvDiffType
    {
        ColumnAdded,
        ColumnRemoved,
        RowAdded,
        RowRemoved,
        CellModified,
        DuplicateKey
    }

    public sealed class CsvDiffItem
    {
        public required CsvDiffType DiffType { get; init; }
        public required string DiffTypeText { get; init; }
        public required string Locator { get; init; }
        public required string LocatorKey { get; init; }
        public string ColumnName { get; init; } = string.Empty;
        public string LeftValue { get; init; } = string.Empty;
        public string RightValue { get; init; } = string.Empty;
        public string Message { get; init; } = string.Empty;
        public int? RowNumber { get; init; }
        public int GroupSortOrder { get; init; }
        public string LeftRowPreview { get; init; } = string.Empty;
        public string RightRowPreview { get; init; } = string.Empty;
    }

    public sealed class CsvCompareResult
    {
        public required IReadOnlyList<CsvDiffItem> DiffItems { get; init; }
        public required IReadOnlyList<string> ValidationErrors { get; init; }
        public int ColumnAddedCount { get; init; }
        public int ColumnRemovedCount { get; init; }
        public int RowAddedCount { get; init; }
        public int RowRemovedCount { get; init; }
        public int CellModifiedCount { get; init; }
        public int DuplicateKeyCount { get; init; }
        public bool HasValidationErrors => ValidationErrors.Count > 0;
    }

    public static class CsvCompareService
    {
        public static CsvCompareResult Compare(
            DataTable leftTable,
            DataTable rightTable,
            CsvCompareMode mode,
            IReadOnlyList<string>? keyColumns = null)
        {
            var diffItems = new List<CsvDiffItem>();
            var validationErrors = new List<string>();

            var commonColumns = BuildCommonColumns(leftTable, rightTable);
            AppendHeaderDifferences(leftTable, rightTable, diffItems);

            if (mode == CsvCompareMode.ByKeyColumns)
            {
                var normalizedKeyColumns = (keyColumns ?? [])
                    .Where(column => !string.IsNullOrWhiteSpace(column))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();

                if (normalizedKeyColumns.Count > 0)
                {
                    CompareByKeyColumns(leftTable, rightTable, commonColumns, normalizedKeyColumns, diffItems, validationErrors);
                }
            }
            else
            {
                CompareByRowNumber(leftTable, rightTable, commonColumns, diffItems);
            }

            return BuildResult(diffItems, validationErrors);
        }

        private static void AppendHeaderDifferences(DataTable leftTable, DataTable rightTable, List<CsvDiffItem> diffItems)
        {
            var rightColumns = new HashSet<string>(rightTable.Columns.Cast<DataColumn>().Select(column => column.ColumnName), StringComparer.OrdinalIgnoreCase);
            var leftColumns = new HashSet<string>(leftTable.Columns.Cast<DataColumn>().Select(column => column.ColumnName), StringComparer.OrdinalIgnoreCase);

            foreach (DataColumn column in leftTable.Columns)
            {
                if (!rightColumns.Contains(column.ColumnName))
                {
                    diffItems.Add(CreateDiffItem(
                        CsvDiffType.ColumnRemoved,
                        "\u8868\u5934",
                        "HEADER",
                        column.ColumnName,
                        column.ColumnName,
                        string.Empty,
                        $"\u8be5\u5b57\u6bb5\u4ec5\u5b58\u5728\u4e8e A \u8868\uff0cB \u8868\u7f3a\u5931\uff1a{column.ColumnName}",
                        null,
                        string.Empty,
                        string.Empty));
                }
            }

            foreach (DataColumn column in rightTable.Columns)
            {
                if (!leftColumns.Contains(column.ColumnName))
                {
                    diffItems.Add(CreateDiffItem(
                        CsvDiffType.ColumnAdded,
                        "\u8868\u5934",
                        "HEADER",
                        column.ColumnName,
                        string.Empty,
                        column.ColumnName,
                        $"\u8be5\u5b57\u6bb5\u4ec5\u5b58\u5728\u4e8e B \u8868\uff0cA \u8868\u7f3a\u5931\uff1a{column.ColumnName}",
                        null,
                        string.Empty,
                        string.Empty));
                }
            }
        }

        private static List<string> BuildCommonColumns(DataTable leftTable, DataTable rightTable)
        {
            var rightColumns = new HashSet<string>(rightTable.Columns.Cast<DataColumn>().Select(column => column.ColumnName), StringComparer.OrdinalIgnoreCase);
            return leftTable.Columns.Cast<DataColumn>()
                .Select(column => column.ColumnName)
                .Where(rightColumns.Contains)
                .ToList();
        }

        private static void CompareByRowNumber(
            DataTable leftTable,
            DataTable rightTable,
            IReadOnlyList<string> commonColumns,
            List<CsvDiffItem> diffItems)
        {
            int maxRowCount = Math.Max(leftTable.Rows.Count, rightTable.Rows.Count);

            for (int rowIndex = 0; rowIndex < maxRowCount; rowIndex++)
            {
                int rowNumber = rowIndex + 1;
                bool hasLeftRow = rowIndex < leftTable.Rows.Count;
                bool hasRightRow = rowIndex < rightTable.Rows.Count;

                if (hasLeftRow && !hasRightRow)
                {
                    DataRow leftRow = leftTable.Rows[rowIndex];
                    diffItems.Add(CreateDiffItem(
                        CsvDiffType.RowRemoved,
                        $"\u7b2c {rowNumber} \u884c",
                        $"ROW:{rowNumber}",
                        string.Empty,
                        BuildRowPreview(leftRow, leftTable.Columns),
                        string.Empty,
                        "\u8be5\u884c\u4ec5\u5b58\u5728\u4e8e A \u8868\uff0cB \u8868\u7f3a\u5931",
                        rowNumber,
                        BuildRowPreview(leftRow, leftTable.Columns),
                        string.Empty));
                    continue;
                }

                if (!hasLeftRow && hasRightRow)
                {
                    DataRow rightRow = rightTable.Rows[rowIndex];
                    diffItems.Add(CreateDiffItem(
                        CsvDiffType.RowAdded,
                        $"\u7b2c {rowNumber} \u884c",
                        $"ROW:{rowNumber}",
                        string.Empty,
                        string.Empty,
                        BuildRowPreview(rightRow, rightTable.Columns),
                        "\u8be5\u884c\u4ec5\u5b58\u5728\u4e8e B \u8868\uff0cA \u8868\u7f3a\u5931",
                        rowNumber,
                        string.Empty,
                        BuildRowPreview(rightRow, rightTable.Columns)));
                    continue;
                }

                DataRow leftDataRow = leftTable.Rows[rowIndex];
                DataRow rightDataRow = rightTable.Rows[rowIndex];

                foreach (string columnName in commonColumns)
                {
                    string leftValue = GetCellText(leftDataRow, columnName);
                    string rightValue = GetCellText(rightDataRow, columnName);

                    if (!string.Equals(leftValue, rightValue, StringComparison.Ordinal))
                    {
                        diffItems.Add(CreateDiffItem(
                            CsvDiffType.CellModified,
                            $"\u7b2c {rowNumber} \u884c",
                            $"ROW:{rowNumber}",
                            columnName,
                            leftValue,
                            rightValue,
                            "\u540c\u4e00\u884c\u540c\u4e00\u5217\u7684\u503c\u4e0d\u4e00\u81f4",
                            rowNumber,
                            BuildRowPreview(leftDataRow, leftTable.Columns),
                            BuildRowPreview(rightDataRow, rightTable.Columns)));
                    }
                }
            }
        }

        private static void CompareByKeyColumns(
            DataTable leftTable,
            DataTable rightTable,
            IReadOnlyList<string> commonColumns,
            IReadOnlyList<string> keyColumns,
            List<CsvDiffItem> diffItems,
            List<string> validationErrors)
        {
            var comparableColumns = commonColumns
                .Where(column => !keyColumns.Contains(column, StringComparer.OrdinalIgnoreCase))
                .ToList();

            var leftLookup = BuildKeyLookup(leftTable, keyColumns, validationErrors, "A", diffItems);
            var rightLookup = BuildKeyLookup(rightTable, keyColumns, validationErrors, "B", diffItems);

            var allKeys = leftLookup.Keys
                .Concat(rightLookup.Keys)
                .Distinct(StringComparer.Ordinal)
                .OrderBy(key => key, StringComparer.Ordinal)
                .ToList();

            foreach (string key in allKeys)
            {
                bool hasLeft = leftLookup.TryGetValue(key, out var leftRows);
                bool hasRight = rightLookup.TryGetValue(key, out var rightRows);

                if (hasLeft && !hasRight && leftRows != null)
                {
                    var leftEntry = leftRows[0];
                    diffItems.Add(CreateDiffItem(
                        CsvDiffType.RowRemoved,
                        BuildKeyLocator(keyColumns, leftEntry.Row),
                        $"KEY:{key}",
                        string.Empty,
                        BuildRowPreview(leftEntry.Row, leftTable.Columns),
                        string.Empty,
                        "\u8be5\u4e3b\u952e\u4ec5\u5b58\u5728\u4e8e A \u8868\uff0cB \u8868\u7f3a\u5931",
                        leftEntry.RowNumber,
                        BuildRowPreview(leftEntry.Row, leftTable.Columns),
                        string.Empty));
                    continue;
                }

                if (!hasLeft && hasRight && rightRows != null)
                {
                    var rightEntry = rightRows[0];
                    diffItems.Add(CreateDiffItem(
                        CsvDiffType.RowAdded,
                        BuildKeyLocator(keyColumns, rightEntry.Row),
                        $"KEY:{key}",
                        string.Empty,
                        string.Empty,
                        BuildRowPreview(rightEntry.Row, rightTable.Columns),
                        "\u8be5\u4e3b\u952e\u4ec5\u5b58\u5728\u4e8e B \u8868\uff0cA \u8868\u7f3a\u5931",
                        rightEntry.RowNumber,
                        string.Empty,
                        BuildRowPreview(rightEntry.Row, rightTable.Columns)));
                    continue;
                }

                if (leftRows == null || rightRows == null || leftRows.Count != 1 || rightRows.Count != 1)
                {
                    continue;
                }

                var leftItem = leftRows[0];
                var rightItem = rightRows[0];
                string locator = BuildKeyLocator(keyColumns, leftItem.Row);
                string locatorKey = $"KEY:{key}";

                foreach (string columnName in comparableColumns)
                {
                    string leftValue = GetCellText(leftItem.Row, columnName);
                    string rightValue = GetCellText(rightItem.Row, columnName);

                    if (!string.Equals(leftValue, rightValue, StringComparison.Ordinal))
                    {
                        diffItems.Add(CreateDiffItem(
                            CsvDiffType.CellModified,
                            locator,
                            locatorKey,
                            columnName,
                            leftValue,
                            rightValue,
                            "\u76f8\u540c\u4e3b\u952e\u4e0b\u8be5\u5217\u7684\u503c\u4e0d\u4e00\u81f4",
                            leftItem.RowNumber,
                            BuildRowPreview(leftItem.Row, leftTable.Columns),
                            BuildRowPreview(rightItem.Row, rightTable.Columns)));
                    }
                }
            }
        }

        private static Dictionary<string, List<RowEntry>> BuildKeyLookup(
            DataTable table,
            IReadOnlyList<string> keyColumns,
            List<string> validationErrors,
            string sideLabel,
            List<CsvDiffItem> diffItems)
        {
            var lookup = new Dictionary<string, List<RowEntry>>(StringComparer.Ordinal);

            for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
            {
                DataRow row = table.Rows[rowIndex];
                string key = BuildCompositeKey(keyColumns, row);

                if (!lookup.TryGetValue(key, out var entries))
                {
                    entries = [];
                    lookup[key] = entries;
                }

                entries.Add(new RowEntry(row, rowIndex + 1));
            }

            foreach ((string key, List<RowEntry> entries) in lookup.Where(pair => pair.Value.Count > 1))
            {
                string duplicatedRows = string.Join("\u3001", entries.Select(entry => entry.RowNumber.ToString(CultureInfo.InvariantCulture)));
                string message = $"{sideLabel} \u8868\u5b58\u5728\u91cd\u590d\u4e3b\u952e\uff1a{FormatLocatorValue(key)}\uff0c\u91cd\u590d\u884c\u53f7\uff1a{duplicatedRows}";
                validationErrors.Add(message);

                foreach (RowEntry entry in entries)
                {
                    diffItems.Add(CreateDiffItem(
                        CsvDiffType.DuplicateKey,
                        BuildKeyLocator(keyColumns, entry.Row),
                        $"DUP:{sideLabel}:{key}",
                        string.Empty,
                        sideLabel == "A" ? BuildRowPreview(entry.Row, table.Columns) : string.Empty,
                        sideLabel == "B" ? BuildRowPreview(entry.Row, table.Columns) : string.Empty,
                        message,
                        entry.RowNumber,
                        sideLabel == "A" ? BuildRowPreview(entry.Row, table.Columns) : string.Empty,
                        sideLabel == "B" ? BuildRowPreview(entry.Row, table.Columns) : string.Empty));
                }
            }

            return lookup;
        }

        private static string BuildCompositeKey(IReadOnlyList<string> keyColumns, DataRow row)
        {
            var parts = keyColumns
                .Select(column => $"{column}={GetCellText(row, column)}")
                .ToArray();

            return string.Join("\u001f", parts);
        }

        private static string BuildKeyLocator(IReadOnlyList<string> keyColumns, DataRow row)
        {
            var parts = keyColumns
                .Select(column => $"{column}={FormatLocatorValue(GetCellText(row, column))}")
                .ToArray();

            return string.Join("\u3001", parts);
        }

        private static string FormatLocatorValue(string value)
        {
            return string.IsNullOrWhiteSpace(value) ? "(\u7a7a)" : value;
        }

        private static string BuildRowPreview(DataRow row, DataColumnCollection columns)
        {
            var parts = new List<string>();

            for (int index = 0; index < columns.Count && index < 4; index++)
            {
                string columnName = columns[index].ColumnName;
                parts.Add($"{columnName}={FormatPreviewValue(GetCellText(row, columnName))}");
            }

            return string.Join(" | ", parts);
        }

        private static string FormatPreviewValue(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return "(\u7a7a)";
            }

            if (value.Length <= 28)
            {
                return value;
            }

            return $"{value[..25]}...";
        }

        private static string GetCellText(DataRow row, string columnName)
        {
            if (!row.Table.Columns.Contains(columnName))
            {
                return string.Empty;
            }

            object value = row[columnName];
            return value == DBNull.Value ? string.Empty : Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
        }

        private static CsvDiffItem CreateDiffItem(
            CsvDiffType diffType,
            string locator,
            string locatorKey,
            string columnName,
            string leftValue,
            string rightValue,
            string message,
            int? rowNumber,
            string leftRowPreview,
            string rightRowPreview)
        {
            return new CsvDiffItem
            {
                DiffType = diffType,
                DiffTypeText = GetDiffTypeText(diffType),
                Locator = locator,
                LocatorKey = locatorKey,
                ColumnName = columnName,
                LeftValue = leftValue,
                RightValue = rightValue,
                Message = message,
                RowNumber = rowNumber,
                GroupSortOrder = GetGroupSortOrder(diffType),
                LeftRowPreview = leftRowPreview,
                RightRowPreview = rightRowPreview
            };
        }

        private static int GetGroupSortOrder(CsvDiffType diffType)
        {
            return diffType switch
            {
                CsvDiffType.ColumnRemoved => 1,
                CsvDiffType.ColumnAdded => 2,
                CsvDiffType.DuplicateKey => 3,
                CsvDiffType.RowRemoved => 4,
                CsvDiffType.RowAdded => 5,
                CsvDiffType.CellModified => 6,
                _ => 99
            };
        }

        private static string GetDiffTypeText(CsvDiffType diffType)
        {
            return diffType switch
            {
                CsvDiffType.ColumnAdded => "\u5b57\u6bb5\u65b0\u589e",
                CsvDiffType.ColumnRemoved => "\u5b57\u6bb5\u7f3a\u5931",
                CsvDiffType.RowAdded => "\u884c\u65b0\u589e",
                CsvDiffType.RowRemoved => "\u884c\u7f3a\u5931",
                CsvDiffType.CellModified => "\u5355\u5143\u683c\u5dee\u5f02",
                CsvDiffType.DuplicateKey => "\u4e3b\u952e\u91cd\u590d",
                _ => string.Empty
            };
        }

        private static CsvCompareResult BuildResult(List<CsvDiffItem> diffItems, List<string> validationErrors)
        {
            return new CsvCompareResult
            {
                DiffItems = diffItems,
                ValidationErrors = validationErrors,
                ColumnAddedCount = diffItems.Count(item => item.DiffType == CsvDiffType.ColumnAdded),
                ColumnRemovedCount = diffItems.Count(item => item.DiffType == CsvDiffType.ColumnRemoved),
                RowAddedCount = diffItems.Count(item => item.DiffType == CsvDiffType.RowAdded),
                RowRemovedCount = diffItems.Count(item => item.DiffType == CsvDiffType.RowRemoved),
                CellModifiedCount = diffItems.Count(item => item.DiffType == CsvDiffType.CellModified),
                DuplicateKeyCount = diffItems.Count(item => item.DiffType == CsvDiffType.DuplicateKey)
            };
        }

        private readonly record struct RowEntry(DataRow Row, int RowNumber);
    }
}
