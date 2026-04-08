using System.Data;

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
        public CsvDiffType DiffType { get; init; }

        public string DiffTypeText { get; init; } = "";

        public string Locator { get; init; } = "";

        public string ColumnName { get; init; } = "";

        public string LeftValue { get; init; } = "";

        public string RightValue { get; init; } = "";

        public string Message { get; init; } = "";
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
        private sealed class ColumnPair
        {
            public required string DisplayName { get; init; }

            public required string LeftName { get; init; }

            public required string RightName { get; init; }
        }

        public static CsvCompareResult Compare(
            DataTable leftTable,
            DataTable rightTable,
            CsvCompareMode mode,
            IReadOnlyList<string>? keyColumns = null)
        {
            var diffItems = new List<CsvDiffItem>();
            var validationErrors = new List<string>();
            var leftColumns = leftTable.Columns.Cast<DataColumn>().Select(column => column.ColumnName).ToList();
            var rightColumns = rightTable.Columns.Cast<DataColumn>().Select(column => column.ColumnName).ToList();
            var commonColumns = BuildCommonColumns(leftColumns, rightColumns);

            var leftLookup = leftColumns.ToHashSet(StringComparer.OrdinalIgnoreCase);
            var rightLookup = rightColumns.ToHashSet(StringComparer.OrdinalIgnoreCase);

            foreach (string leftOnlyColumn in leftColumns.Where(column => !rightLookup.Contains(column)))
            {
                diffItems.Add(CreateDiffItem(
                    CsvDiffType.ColumnRemoved,
                    "表头",
                    leftOnlyColumn,
                    leftOnlyColumn,
                    "",
                    "列只存在于 A 文件"));
            }

            foreach (string rightOnlyColumn in rightColumns.Where(column => !leftLookup.Contains(column)))
            {
                diffItems.Add(CreateDiffItem(
                    CsvDiffType.ColumnAdded,
                    "表头",
                    rightOnlyColumn,
                    "",
                    rightOnlyColumn,
                    "列只存在于 B 文件"));
            }

            if (mode == CsvCompareMode.ByKeyColumns)
            {
                CompareByKeyColumns(leftTable, rightTable, commonColumns, keyColumns ?? [], diffItems, validationErrors);
            }
            else
            {
                CompareByRowNumber(leftTable, rightTable, commonColumns, diffItems);
            }

            return BuildResult(diffItems, validationErrors);
        }

        private static void CompareByRowNumber(
            DataTable leftTable,
            DataTable rightTable,
            IReadOnlyList<ColumnPair> commonColumns,
            List<CsvDiffItem> diffItems)
        {
            int maxRows = Math.Max(leftTable.Rows.Count, rightTable.Rows.Count);

            for (int rowIndex = 0; rowIndex < maxRows; rowIndex++)
            {
                string locator = $"第 {rowIndex + 1} 行";

                if (rowIndex >= leftTable.Rows.Count)
                {
                    diffItems.Add(CreateDiffItem(
                        CsvDiffType.RowAdded,
                        locator,
                        "",
                        "",
                        "",
                        "该行只存在于 B 文件"));
                    continue;
                }

                if (rowIndex >= rightTable.Rows.Count)
                {
                    diffItems.Add(CreateDiffItem(
                        CsvDiffType.RowRemoved,
                        locator,
                        "",
                        "",
                        "",
                        "该行只存在于 A 文件"));
                    continue;
                }

                DataRow leftRow = leftTable.Rows[rowIndex];
                DataRow rightRow = rightTable.Rows[rowIndex];

                foreach (ColumnPair columnPair in commonColumns)
                {
                    string leftValue = GetCellText(leftRow, columnPair.LeftName);
                    string rightValue = GetCellText(rightRow, columnPair.RightName);
                    if (string.Equals(leftValue, rightValue, StringComparison.Ordinal))
                    {
                        continue;
                    }

                    diffItems.Add(CreateDiffItem(
                        CsvDiffType.CellModified,
                        locator,
                        columnPair.DisplayName,
                        leftValue,
                        rightValue,
                        "同一行同一列的值不同"));
                }
            }
        }

        private static void CompareByKeyColumns(
            DataTable leftTable,
            DataTable rightTable,
            IReadOnlyList<ColumnPair> commonColumns,
            IReadOnlyList<string> selectedKeyColumns,
            List<CsvDiffItem> diffItems,
            List<string> validationErrors)
        {
            var selectedPairs = commonColumns
                .Where(pair => selectedKeyColumns.Contains(pair.DisplayName, StringComparer.OrdinalIgnoreCase))
                .ToList();

            if (selectedPairs.Count == 0)
            {
                validationErrors.Add("主键模式至少需要选择一个公共列。");
                return;
            }

            Dictionary<string, DataRow> leftRows = BuildKeyLookup(leftTable, selectedPairs, true, diffItems, validationErrors);
            Dictionary<string, DataRow> rightRows = BuildKeyLookup(rightTable, selectedPairs, false, diffItems, validationErrors);

            if (validationErrors.Count > 0)
            {
                return;
            }

            var leftKeys = leftRows.Keys.ToHashSet(StringComparer.Ordinal);
            var rightKeys = rightRows.Keys.ToHashSet(StringComparer.Ordinal);

            foreach (string removedKey in leftKeys.Except(rightKeys, StringComparer.Ordinal))
            {
                DataRow row = leftRows[removedKey];
                diffItems.Add(CreateDiffItem(
                    CsvDiffType.RowRemoved,
                    BuildKeyLocator(row, selectedPairs, true),
                    "",
                    "",
                    "",
                    "该主键只存在于 A 文件"));
            }

            foreach (string addedKey in rightKeys.Except(leftKeys, StringComparer.Ordinal))
            {
                DataRow row = rightRows[addedKey];
                diffItems.Add(CreateDiffItem(
                    CsvDiffType.RowAdded,
                    BuildKeyLocator(row, selectedPairs, false),
                    "",
                    "",
                    "",
                    "该主键只存在于 B 文件"));
            }

            var keyDisplayNames = selectedPairs
                .Select(pair => pair.DisplayName)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            foreach (string commonKey in leftKeys.Intersect(rightKeys, StringComparer.Ordinal))
            {
                DataRow leftRow = leftRows[commonKey];
                DataRow rightRow = rightRows[commonKey];

                foreach (ColumnPair columnPair in commonColumns)
                {
                    if (keyDisplayNames.Contains(columnPair.DisplayName))
                    {
                        continue;
                    }

                    string leftValue = GetCellText(leftRow, columnPair.LeftName);
                    string rightValue = GetCellText(rightRow, columnPair.RightName);
                    if (string.Equals(leftValue, rightValue, StringComparison.Ordinal))
                    {
                        continue;
                    }

                    diffItems.Add(CreateDiffItem(
                        CsvDiffType.CellModified,
                        BuildKeyLocator(leftRow, selectedPairs, true),
                        columnPair.DisplayName,
                        leftValue,
                        rightValue,
                        "相同主键下该列的值不同"));
                }
            }
        }

        private static Dictionary<string, DataRow> BuildKeyLookup(
            DataTable table,
            IReadOnlyList<ColumnPair> keyColumns,
            bool isLeftSide,
            List<CsvDiffItem> diffItems,
            List<string> validationErrors)
        {
            var rowsByKey = new Dictionary<string, List<int>>(StringComparer.Ordinal);
            var firstRowsByKey = new Dictionary<string, DataRow>(StringComparer.Ordinal);

            for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
            {
                DataRow row = table.Rows[rowIndex];
                string compositeKey = BuildCompositeKey(row, keyColumns, isLeftSide);

                if (!rowsByKey.TryGetValue(compositeKey, out List<int>? indexes))
                {
                    indexes = [];
                    rowsByKey[compositeKey] = indexes;
                    firstRowsByKey[compositeKey] = row;
                }

                indexes.Add(rowIndex + 1);
            }

            foreach ((string compositeKey, List<int> rowIndexes) in rowsByKey.Where(item => item.Value.Count > 1))
            {
                DataRow sampleRow = firstRowsByKey[compositeKey];
                string sideLabel = isLeftSide ? "A" : "B";
                string locator = BuildKeyLocator(sampleRow, keyColumns, isLeftSide);
                string rowsText = string.Join(", ", rowIndexes);

                validationErrors.Add($"{sideLabel} 文件存在重复主键：{locator}，重复行号：{rowsText}");
                diffItems.Add(CreateDiffItem(
                    CsvDiffType.DuplicateKey,
                    locator,
                    "",
                    isLeftSide ? rowsText : "",
                    isLeftSide ? "" : rowsText,
                    $"{sideLabel} 文件存在重复主键"));
            }

            if (validationErrors.Count > 0)
            {
                return new Dictionary<string, DataRow>(StringComparer.Ordinal);
            }

            return firstRowsByKey;
        }

        private static IReadOnlyList<ColumnPair> BuildCommonColumns(
            IReadOnlyList<string> leftColumns,
            IReadOnlyList<string> rightColumns)
        {
            var rightColumnsByName = rightColumns.ToDictionary(column => column, StringComparer.OrdinalIgnoreCase);
            var columnPairs = new List<ColumnPair>();

            foreach (string leftColumn in leftColumns)
            {
                if (!rightColumnsByName.TryGetValue(leftColumn, out string? rightColumn))
                {
                    continue;
                }

                columnPairs.Add(new ColumnPair
                {
                    DisplayName = leftColumn,
                    LeftName = leftColumn,
                    RightName = rightColumn
                });
            }

            return columnPairs;
        }

        private static string BuildCompositeKey(DataRow row, IReadOnlyList<ColumnPair> keyColumns, bool useLeftColumns)
        {
            return string.Join("\u001F", keyColumns.Select(pair => GetCellText(row, useLeftColumns ? pair.LeftName : pair.RightName)));
        }

        private static string BuildKeyLocator(DataRow row, IReadOnlyList<ColumnPair> keyColumns, bool useLeftColumns)
        {
            return string.Join(" | ", keyColumns.Select(pair =>
            {
                string value = GetCellText(row, useLeftColumns ? pair.LeftName : pair.RightName);
                return $"{pair.DisplayName}={FormatLocatorValue(value)}";
            }));
        }

        private static string FormatLocatorValue(string value)
            => string.IsNullOrEmpty(value) ? "(空)" : value;

        private static string GetCellText(DataRow row, string columnName)
        {
            object? value = row[columnName];
            return value == DBNull.Value ? string.Empty : value?.ToString() ?? string.Empty;
        }

        private static CsvDiffItem CreateDiffItem(
            CsvDiffType diffType,
            string locator,
            string columnName,
            string leftValue,
            string rightValue,
            string message)
        {
            return new CsvDiffItem
            {
                DiffType = diffType,
                DiffTypeText = GetDiffTypeText(diffType),
                Locator = locator,
                ColumnName = columnName,
                LeftValue = leftValue,
                RightValue = rightValue,
                Message = message
            };
        }

        private static string GetDiffTypeText(CsvDiffType diffType) => diffType switch
        {
            CsvDiffType.ColumnAdded => "列新增",
            CsvDiffType.ColumnRemoved => "列删除",
            CsvDiffType.RowAdded => "行新增",
            CsvDiffType.RowRemoved => "行删除",
            CsvDiffType.CellModified => "单元格修改",
            CsvDiffType.DuplicateKey => "主键重复",
            _ => "未知差异"
        };

        private static CsvCompareResult BuildResult(IReadOnlyList<CsvDiffItem> diffItems, IReadOnlyList<string> validationErrors)
        {
            return new CsvCompareResult
            {
                DiffItems = diffItems.ToList(),
                ValidationErrors = validationErrors.ToList(),
                ColumnAddedCount = diffItems.Count(item => item.DiffType == CsvDiffType.ColumnAdded),
                ColumnRemovedCount = diffItems.Count(item => item.DiffType == CsvDiffType.ColumnRemoved),
                RowAddedCount = diffItems.Count(item => item.DiffType == CsvDiffType.RowAdded),
                RowRemovedCount = diffItems.Count(item => item.DiffType == CsvDiffType.RowRemoved),
                CellModifiedCount = diffItems.Count(item => item.DiffType == CsvDiffType.CellModified),
                DuplicateKeyCount = diffItems.Count(item => item.DiffType == CsvDiffType.DuplicateKey)
            };
        }
    }
}
