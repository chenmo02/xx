using System.Data;

namespace WpfApp1.Services
{
    public sealed class CsvCompareCachedFile
    {
        public required string FilePath { get; init; }
        public required string FileName { get; init; }
        public required long FileSize { get; init; }
        public required string Delimiter { get; init; }
        public required System.Text.Encoding Encoding { get; init; }
        public required DataTable Table { get; init; }
    }

    public sealed class CsvCompareCachedState
    {
        public CsvCompareCachedFile? LeftFile { get; set; }
        public CsvCompareCachedFile? RightFile { get; set; }
        public CsvCompareResult? LastResult { get; set; }
        public CsvCompareMode CurrentMode { get; set; } = CsvCompareMode.ByRowNumber;
        public List<string> SelectedKeyColumns { get; set; } = [];
        public bool ShowColumnChanges { get; set; } = true;
        public bool ShowRowChanges { get; set; } = true;
        public bool ShowCellChanges { get; set; } = true;
        public bool ExportFilteredOnly { get; set; }
        public int PageSize { get; set; } = 100;
        public int PageIndex { get; set; }
    }

    public static class CsvCompareStateService
    {
        public static CsvCompareCachedState State { get; } = new();

        public static void Reset()
        {
            State.LeftFile = null;
            State.RightFile = null;
            State.LastResult = null;
            State.CurrentMode = CsvCompareMode.ByRowNumber;
            State.SelectedKeyColumns = [];
            State.ShowColumnChanges = true;
            State.ShowRowChanges = true;
            State.ShowCellChanges = true;
            State.ExportFilteredOnly = false;
            State.PageSize = 100;
            State.PageIndex = 0;
        }
    }
}
