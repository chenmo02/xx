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

        public string LocatorKey { get; init; } = "";

        public string ColumnName { get; init; } = "";

        public string LeftValue { get; init; } = "";

        public string RightValue { get; init; } = "";

        public string Message { get; init; } = "";

        public int? RowNumber { get; init; }

        public int GroupSortOrder { get; init; }

        public string LeftRowPreview { get; init; } = "";

        public string RightRowPreview { get; init; } = "";
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
                    $"HEADER:{leftOnlyColumn}",
                    leftOnlyColumn,
                    leftOnlyColumn,
                    "",
                    "该列仅存在于文件 A",
                    groupSortOrder: GetGroupSortOrder(CsvDiffType.ColumnRemoved)));
            }

            foreach (string rightOnlyColumn in rightColumns.Where(column => !leftLookup.Contains(column)))
            {
                diffItems.Add(CreateDiffItem(
                    CsvDiffType.ColumnAdded,
                    "表头",
                    $"HEADER:{rightOnlyColumn}",
                    rightOnlyColumn,
                    "",
                    rightOnlyColumn,
                    "该列仅存在于文件 B",
                    groupSortOrder: GetGroupSortOrder(CsvDiffType.ColumnAdded)));
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
                int rowNumber = rowIndex + 1;
                string locator = $"第 {rowNumber} 行";
                string locatorKey = $"ROW:{rowNumber}";

                if (rowIndex >= leftTable.Rows.Count)
                {
                    DataRow rightRow = rightTable.Rows[rowIndex];
                    diffItems.Add(CreateDiffItem(
                        CsvDiffType.RowAdded,
                        locator,
                        locatorKey,
                        "",
                        "",
                        "",
                        "该行仅存在于文件 B",
                        rowNumber,
                        GetGroupSortOrder(CsvDiffType.RowAdded),
                        rightRowPreview: BuildRowPreview(rightRow, commonColumns, useLeftColumns: false)));
                    continue;
                }

                if (rowIndex >= rightTable.Rows.Count)
                {
                    DataRow leftRow = leftTable.Rows[rowIndex];
                    diffItems.Add(CreateDiffItem(
                        CsvDiffType.RowRemoved,
                        locator,
                        locatorKey,
                        "",
                        "",
                        "",
                        "该行仅存在于文件 A",
                        rowNumber,
                        GetGroupSortOrder(CsvDiffType.RowRemoved),
                        leftRowPreview: BuildRowPreview(leftRow, commonColumns, useLeftColumns: true)));
                    continue;
                }

                DataRow currentLeftRow = leftTable.Rows[rowIndex];
                DataRow currentRightRow = rightTable.Rows[rowIndex];
                string leftPreview = BuildRowPreview(currentLeftRow, commonColumns, useLeftColumns: true);
                string rightPreview = BuildRowPreview(currentRightRow, commonColumns, useLeftColumns: false);

                foreach (ColumnPair columnPair in commonColumns)
                {
                    string leftValue = GetCellText(currentLeftRow, columnPair.LeftName);
                    string rightValue = GetCellText(currentRightRow, columnPair.RightName);
                    if (string.Equals(leftValue, rightValue, StringComparison.Ordinal))
                    {
                        continue;
                    }

                    diffItems.Add(CreateDiffItem(
                        CsvDiffType.CellModified,
                        locator,
                        locatorKey,
                        columnPair.DisplayName,
                        leftValue,
                        rightValue,
                        "同一行同一列的值不同",
                        rowNumber,
                        GetGroupSortOrder(CsvDiffType.CellModified),
                        leftPreview,
                        rightPreview));
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

            Dictionary<string, DataRow> leftRows = BuildKeyLookup(leftTable, commonColumns, selectedPairs, true, diffItems, validationErrors);
            Dictionary<string, DataRow> rightRows = BuildKeyLookup(rightTable, commonColumns, selectedPairs, false, diffItems, validationErrors);

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
                    BuildKeyLocator(row, selectedPairs, useLeftColumns: true),
                    $"KEY:{removedKey}",
                    "",
                    "",
                    "",
                    "该主键仅存在于文件 A",
                    groupSortOrder: GetGroupSortOrder(CsvDiffType.RowRemoved),
                    leftRowPreview: BuildRowPreview(row, commonColumns, useLeftColumns: true)));
            }

            foreach (string addedKey in rightKeys.Except(leftKeys, StringComparer.Ordinal))
            {
                DataRow row = rightRows[addedKey];
                diffItems.Add(CreateDiffItem(
                    CsvDiffType.RowAdded,
                    BuildKeyLocator(row, selectedPairs, useLeftColumns: false),
                    $"KEY:{addedKey}",
                    "",
                    "",
                    "",
                    "该主键仅存在于文件 B",
                    groupSortOrder: GetGroupSortOrder(CsvDiffType.RowAdded),
                    rightRowPreview: BuildRowPreview(row, commonColumns, useLeftColumns: false)));
            }

            var keyDisplayNames = selectedPairs
                .Select(pair => pair.DisplayName)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            foreach (string commonKey in leftKeys.Intersect(rightKeys, StringComparer.Ordinal))
            {
                DataRow leftRow = leftRows[commonKey];
                DataRow rightRow = rightRows[commonKey];
                string locator = BuildKeyLocator(leftRow, selectedPairs, useLeftColumns: true);
                string leftPreview = BuildRowPreview(leftRow, commonColumns, useLeftColumns: true);
                string rightPreview = BuildRowPreview(rightRow, commonColumns, useLeftColumns: false);

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
                        locator,
                        $"KEY:{commonKey}",
                        columnPair.DisplayName,
                        leftValue,
                        rightValue,
                        "相同主键下该列的值不同",
                        groupSortOrder: GetGroupSortOrder(CsvDiffType.CellModified),
                        leftRowPreview: leftPreview,
                        rightRowPreview: rightPreview));
                }
            }
        }

        private static Dictionary<string, DataRow> BuildKeyLookup(
            DataTable table,
            IReadOnlyList<ColumnPair> previewColumns,
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
                string leftPreview = isLeftSide ? BuildRowPreview(sampleRow, previewColumns, useLeftColumns: true) : "";
                string rightPreview = isLeftSide ? "" : BuildRowPreview(sampleRow, previewColumns, useLeftColumns: false);

                validationErrors.Add($"{sideLabel} 文件存在重复主键：{locator}，重复行号：{rowsText}");
                diffItems.Add(CreateDiffItem(
                    CsvDiffType.DuplicateKey,
                    locator,
                    $"DUP:{sideLabel}:{compositeKey}",
                    "",
                    isLeftSide ? rowsText : "",
                    isLeftSide ? "" : rowsText,
                    $"{sideLabel} 文件存在重复主键",
                    groupSortOrder: GetGroupSortOrder(CsvDiffType.DuplicateKey),
                    leftRowPreview: leftPreview,
                    rightRowPreview: rightPreview));
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

        private static string BuildRowPreview(DataRow row, IReadOnlyList<ColumnPair> columnPairs, bool useLeftColumns)
        {
            IEnumerable<(string DisplayName, string ColumnName)> previewColumns;
            if (columnPairs.Count > 0)
            {
                previewColumns = columnPairs.Select(pair => (pair.DisplayName, useLeftColumns ? pair.LeftName : pair.RightName));
            }
            else
            {
                previewColumns = row.Table.Columns.Cast<DataColumn>().Select(column => (column.ColumnName, column.ColumnName));
            }

            var pairs = previewColumns.Take(4).ToList();
            if (pairs.Count == 0)
            {
                return "(空行)";
            }

            var segments = new List<string>(pairs.Count);
            foreach ((string displayName, string columnName) in pairs)
            {
                segments.Add($"{displayName}={FormatPreviewValue(GetCellText(row, columnName))}");
            }

            if (previewColumns.Skip(4).Any())
            {
                segments.Add("…");
            }

            return string.Join("；", segments);
        }

        private static string FormatPreviewValue(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return "(空)";
            }

            const int maxLength = 24;
            return value.Length <= maxLength ? value : $"{value[..maxLength]}…";
        }

        private static string GetCellText(DataRow row, string columnName)
        {
            object? value = row[columnName];
            return value == DBNull.Value ? string.Empty : value?.ToString() ?? string.Empty;
        }

        private static CsvDiffItem CreateDiffItem(
            CsvDiffType diffType,
            string locator,
            string locatorKey,
            string columnName,
            string leftValue,
            string rightValue,
            string message,
            int? rowNumber = null,
            int groupSortOrder = 0,
            string leftRowPreview = "",
            string rightRowPreview = "")
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
                GroupSortOrder = groupSortOrder,
                LeftRowPreview = leftRowPreview,
                RightRowPreview = rightRowPreview
            };
        }

        private static int GetGroupSortOrder(CsvDiffType diffType) => diffType switch
        {
            CsvDiffType.DuplicateKey => 0,
            CsvDiffType.ColumnRemoved => 10,
            CsvDiffType.ColumnAdded => 11,
            CsvDiffType.RowRemoved => 20,
            CsvDiffType.RowAdded => 30,
            CsvDiffType.CellModified => 40,
            _ => 99
        };

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
