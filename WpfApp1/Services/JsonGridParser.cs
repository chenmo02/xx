using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Text.Json;

namespace WpfApp1.Services
{
    public class JsonGridNode : INotifyPropertyChanged
    {
        private string _value = "";
        private bool _isExpanded = false;
        private bool _isHighlighted = false;
        public string Key { get; set; } = "";
        public string Value
        {
            get => _value;
            set { _value = value; OnPropertyChanged(); }
        }
        public string NodeType { get; set; } = "";
        /// <summary>JSON 路径（用于搜索联动定位）</summary>
        public string JsonPath { get; set; } = "";
        /// <summary>节点是否展开</summary>
        public bool IsExpanded
        {
            get => _isExpanded;
            set { _isExpanded = value; OnPropertyChanged(); }
        }
        /// <summary>节点是否被搜索高亮</summary>
        public bool IsHighlighted
        {
            get => _isHighlighted;
            set { _isHighlighted = value; OnPropertyChanged(); }
        }
        public ObservableCollection<JsonGridNode> Children { get; set; } = [];
        public bool IsContainer => NodeType is "Object" or "Array";

        public string ExpandLabel => NodeType switch
        {
            "Array" => IsObjectArray
                ? $"[-] {Key}[{DisplayTableRowCount}]"
                : IsSimpleArray
                    ? $"[-] {Key}[{DisplayTableRowCount}]"
                    : $"[-] {Key}[{Children.Count}]",
            "Object" => $"[-] {Key}{{{Children.Count} 项}}",
            _ => Key
        };

        public string ValueColor => NodeType switch
        {
            "String" => "#16A34A",
            "Number" => "#D97706",
            "True" or "False" => "#7C3AED",
            "Null" => "#9CA3AF",
            "Array" => "#2563EB",
            "Object" => "#2563EB",
            _ => "#374151"
        };

        public ObservableCollection<string> TableColumns { get; set; } = [];
        public ObservableCollection<JsonGridRow> TableRows { get; set; } = [];
        public int TableTotalRows { get; set; } = 0;
        public int TableTotalColumns { get; set; } = 0;
        public int DisplayTableRowCount => TableTotalRows > 0 ? TableTotalRows : TableRows.Count;
        public bool IsTableColumnPreviewTruncated { get; set; } = false;
        public bool IsTablePreviewTruncated =>
            TableTotalRows > TableRows.Count || IsTableColumnPreviewTruncated;
        /// <summary>对象数组（每个元素是 object）</summary>
        public bool IsObjectArray { get; set; } = false;
        /// <summary>简单数组（每个元素是标量值）</summary>
        public bool IsSimpleArray { get; set; } = false;
        /// <summary>是否有表格数据可展示</summary>
        public bool HasTable => TableRows.Count > 0;

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string? name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    public class JsonGridRow
    {
        public int Index { get; set; }
        public ObservableCollection<JsonGridCell> Cells { get; set; } = [];
    }

    public class JsonGridCell : INotifyPropertyChanged
    {
        public string ColumnName { get; set; } = "";
        public string JsonPath { get; set; } = "";
        private string _value = "";
        public string Value
        {
            get => _value;
            set { _value = value; OnPropertyChanged(); }
        }
        public string NodeType { get; set; } = "";
        public bool IsNested => NodeType is "Object" or "Array";
        public string NestedSummary { get; set; } = "";
        public ObservableCollection<JsonGridNode> NestedChildren { get; set; } = [];

        public string ValueColor => NodeType switch
        {
            "String" => "#16A34A",
            "Number" => "#D97706",
            "True" or "False" => "#7C3AED",
            "Null" => "#9CA3AF",
            _ => "#374151"
        };

        public event PropertyChangedEventHandler? PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string? name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    public static class JsonGridParser
    {
        private const int MaxTablePreviewRows = 1_000;
        private const int MaxTablePreviewColumns = 80;

        public static ObservableCollection<JsonGridNode> Parse(string json)
        {
            using var doc = JsonDocument.Parse(json);
            var root = ParseElement(doc.RootElement, "root", "");
            return [root];
        }

        private static JsonGridNode ParseElement(JsonElement element, string key, string parentPath)
        {
            var currentPath = string.IsNullOrEmpty(parentPath) ? key : AppendJsonPathSegment(parentPath, key);
            var node = new JsonGridNode
            {
                Key = key,
                NodeType = element.ValueKind.ToString(),
                JsonPath = currentPath
            };

            switch (element.ValueKind)
            {
                case JsonValueKind.Object:
                    node.Value = $"{{{element.EnumerateObject().Count()} 项}}";
                    foreach (var prop in element.EnumerateObject())
                        node.Children.Add(ParseElement(prop.Value, prop.Name, currentPath));
                    break;

                case JsonValueKind.Array:
                    node.Value = $"[{element.GetArrayLength()}]";
                    if (IsObjectArray(element))
                    {
                        node.IsObjectArray = true;
                        BuildObjectArrayTable(element, node, currentPath);
                    }
                    else
                    {
                        // 简单数组：构建为单列表格（直接显示值，不显示 [0],[1],[2]）
                        node.IsSimpleArray = true;
                        BuildSimpleArrayTable(element, node, key, currentPath);
                    }
                    break;

                case JsonValueKind.String:
                    node.Value = element.GetString() ?? "";
                    break;
                case JsonValueKind.Number:
                    node.Value = element.GetRawText();
                    break;
                case JsonValueKind.True:
                    node.Value = "true";
                    break;
                case JsonValueKind.False:
                    node.Value = "false";
                    break;
                case JsonValueKind.Null:
                    node.Value = "null";
                    break;
            }

            return node;
        }

        private static bool IsObjectArray(JsonElement array)
        {
            if (array.GetArrayLength() == 0) return false;
            foreach (var item in array.EnumerateArray())
            {
                if (item.ValueKind != JsonValueKind.Object) return false;
            }
            return true;
        }

        /// <summary>
        /// 简单数组 → 单列表格（如 ["Vue","React","TypeScript"] → 一列 value）
        /// </summary>
        private static void BuildSimpleArrayTable(JsonElement array, JsonGridNode node, string key, string parentPath)
        {
            node.TableTotalRows = array.GetArrayLength();
            node.TableTotalColumns = 1;
            node.TableColumns.Add(key);

            int idx = 0;
            foreach (var item in array.EnumerateArray())
            {
                if (idx >= MaxTablePreviewRows)
                    break;

                idx++;
                var row = new JsonGridRow { Index = idx };
                var cell = new JsonGridCell
                {
                    ColumnName = key,
                    JsonPath = AppendJsonPathSegment(parentPath, (idx - 1).ToString()),
                    NodeType = item.ValueKind.ToString()
                };

                if (item.ValueKind is JsonValueKind.Object or JsonValueKind.Array)
                {
                    // 嵌套
                    cell.NestedSummary = item.ValueKind == JsonValueKind.Array
                        ? $"[+] [{item.GetArrayLength()}]"
                        : $"[+] {{{item.EnumerateObject().Count()} 项}}";
                    cell.NestedChildren.Add(ParseElement(item, (idx - 1).ToString(), parentPath));
                }
                else
                {
                    cell.Value = item.ValueKind switch
                    {
                        JsonValueKind.String => item.GetString() ?? "",
                        JsonValueKind.Number => item.GetRawText(),
                        JsonValueKind.True => "true",
                        JsonValueKind.False => "false",
                        JsonValueKind.Null => "null",
                        _ => item.GetRawText()
                    };
                }

                row.Cells.Add(cell);
                node.TableRows.Add(row);
            }
        }

        /// <summary>
        /// 对象数组 → 多列表格
        /// </summary>
        private static void BuildObjectArrayTable(JsonElement array, JsonGridNode node, string parentPath)
        {
            node.TableTotalRows = array.GetArrayLength();
            var columns = new List<string>();
            var seenColumns = new HashSet<string>(StringComparer.Ordinal);
            int scannedRows = 0;

            foreach (var item in array.EnumerateArray())
            {
                if (scannedRows >= MaxTablePreviewRows)
                    break;

                scannedRows++;
                foreach (var prop in item.EnumerateObject())
                {
                    if (seenColumns.Contains(prop.Name))
                        continue;

                    if (columns.Count >= MaxTablePreviewColumns)
                    {
                        node.IsTableColumnPreviewTruncated = true;
                        continue;
                    }

                    seenColumns.Add(prop.Name);
                    columns.Add(prop.Name);
                }
            }

            node.TableTotalColumns = columns.Count;
            foreach (var col in columns)
                node.TableColumns.Add(col);

            int rowIdx = 0;
            foreach (var item in array.EnumerateArray())
            {
                if (rowIdx >= MaxTablePreviewRows)
                    break;

                rowIdx++;
                var row = new JsonGridRow { Index = rowIdx };

                foreach (var colName in columns)
                {
                    var cell = new JsonGridCell
                    {
                        ColumnName = colName,
                        JsonPath = AppendJsonPathSegment(AppendJsonPathSegment(parentPath, (rowIdx - 1).ToString()), colName)
                    };

                    if (item.TryGetProperty(colName, out var val))
                    {
                        cell.NodeType = val.ValueKind.ToString();

                        if (val.ValueKind is JsonValueKind.Object or JsonValueKind.Array)
                        {
                            cell.Value = "";
                            cell.NestedSummary = val.ValueKind == JsonValueKind.Array
                                ? $"[+] {colName}[{val.GetArrayLength()}]"
                                : $"[+] {{{val.EnumerateObject().Count()} 项}}";
                            cell.NestedChildren.Add(ParseElement(val, colName, AppendJsonPathSegment(parentPath, (rowIdx - 1).ToString())));
                        }
                        else
                        {
                            cell.Value = val.ValueKind switch
                            {
                                JsonValueKind.String => val.GetString() ?? "",
                                JsonValueKind.Number => val.GetRawText(),
                                JsonValueKind.True => "true",
                                JsonValueKind.False => "false",
                                JsonValueKind.Null => "null",
                                _ => val.GetRawText()
                            };
                        }
                    }
                    else
                    {
                        cell.NodeType = "Null";
                        cell.Value = "";
                    }

                    row.Cells.Add(cell);
                }

                node.TableRows.Add(row);
            }
        }

        private static string AppendJsonPathSegment(string parentPath, string segment)
            => $"{parentPath}/{EscapeJsonPathSegment(segment)}";

        private static string EscapeJsonPathSegment(string segment)
            => segment.Replace("~", "~0").Replace("/", "~1");
    }
}
