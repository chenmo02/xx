using System.Data;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using CsvHelper;
using CsvHelper.Configuration;

namespace WpfApp1.Services
{
    /// <summary>
    /// JSON 处理服务
    /// </summary>
    public static class JsonToolService
    {
        // ==================== 格式化 ====================

        /// <summary>
        /// 美化 JSON（缩进格式化）
        /// </summary>
        public static string Beautify(string json)
        {
            using var doc = JsonDocument.Parse(json);
            return JsonSerializer.Serialize(doc.RootElement, new JsonSerializerOptions
            {
                WriteIndented = true,
                Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
            });
        }

        /// <summary>
        /// 压缩 JSON（去除空白）
        /// </summary>
        public static string Minify(string json)
        {
            using var doc = JsonDocument.Parse(json);
            return JsonSerializer.Serialize(doc.RootElement, new JsonSerializerOptions
            {
                WriteIndented = false,
                Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
            });
        }

        // ==================== 校验 ====================

        /// <summary>
        /// 校验 JSON 格式，返回 (是否合法, 错误信息, 错误位置行号)
        /// </summary>
        public static (bool isValid, string message, int errorLine) Validate(string json)
        {
            if (string.IsNullOrWhiteSpace(json))
                return (false, "JSON 内容为空", 0);

            try
            {
                using var doc = JsonDocument.Parse(json, new JsonDocumentOptions
                {
                    AllowTrailingCommas = false,
                    CommentHandling = JsonCommentHandling.Disallow
                });
                
                var root = doc.RootElement;
                int depth = GetMaxDepth(root);
                int nodeCount = CountNodes(root);
                string rootType = root.ValueKind.ToString();

                return (true, $"✅ JSON 格式正确\n类型: {rootType} | 节点数: {nodeCount} | 最大深度: {depth}", 0);
            }
            catch (JsonException ex)
            {
                int line = (int)(ex.LineNumber ?? 0) + 1;
                return (false, $"❌ 第 {line} 行: {ex.Message}", line);
            }
        }

        // ==================== 解析为 DataTable（网格展示） ====================

        /// <summary>
        /// 将 JSON 解析为 DataTable，用于网格展示和编辑
        /// 支持：数组对象 [{...}, {...}] 和单个对象 {...}
        /// </summary>
        public static DataTable ParseToDataTable(string json)
        {
            var dt = new DataTable();
            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;

            if (root.ValueKind == JsonValueKind.Array)
            {
                // 数组：每个元素一行
                // 先收集所有列名
                var columns = new LinkedHashSet<string>();
                foreach (var item in root.EnumerateArray())
                {
                    if (item.ValueKind == JsonValueKind.Object)
                    {
                        foreach (var prop in item.EnumerateObject())
                            columns.Add(prop.Name);
                    }
                }

                if (columns.Count == 0)
                {
                    // 简单数组 [1, 2, 3]
                    dt.Columns.Add("value", typeof(string));
                    foreach (var item in root.EnumerateArray())
                        dt.Rows.Add(GetElementValue(item));
                    return dt;
                }

                foreach (var col in columns)
                    dt.Columns.Add(col, typeof(string));

                foreach (var item in root.EnumerateArray())
                {
                    var row = dt.NewRow();
                    if (item.ValueKind == JsonValueKind.Object)
                    {
                        foreach (var prop in item.EnumerateObject())
                        {
                            if (dt.Columns.Contains(prop.Name))
                                row[prop.Name] = GetElementValue(prop.Value);
                        }
                    }
                    dt.Rows.Add(row);
                }
            }
            else if (root.ValueKind == JsonValueKind.Object)
            {
                // 单个对象：key-value 两列
                dt.Columns.Add("键 (Key)", typeof(string));
                dt.Columns.Add("值 (Value)", typeof(string));
                dt.Columns.Add("类型 (Type)", typeof(string));

                foreach (var prop in root.EnumerateObject())
                {
                    dt.Rows.Add(prop.Name, GetElementValue(prop.Value), prop.Value.ValueKind.ToString());
                }
            }
            else
            {
                dt.Columns.Add("value", typeof(string));
                dt.Rows.Add(GetElementValue(root));
            }

            return dt;
        }

        /// <summary>
        /// 将 DataTable 转回 JSON（网格编辑后同步）
        /// </summary>
        public static string DataTableToJson(DataTable dt, bool isKeyValueMode)
        {
            var options = new JsonSerializerOptions
            {
                WriteIndented = true,
                Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
            };

            if (isKeyValueMode && dt.Columns.Contains("键 (Key)") && dt.Columns.Contains("值 (Value)"))
            {
                // key-value 模式 → 对象
                var dict = new Dictionary<string, object?>();
                foreach (DataRow row in dt.Rows)
                {
                    string key = row["键 (Key)"]?.ToString() ?? "";
                    string val = row["值 (Value)"]?.ToString() ?? "";
                    if (!string.IsNullOrEmpty(key))
                        dict[key] = ParseValue(val);
                }
                return JsonSerializer.Serialize(dict, options);
            }
            else
            {
                // 数组模式 → [{}, {}]
                var list = new List<Dictionary<string, object?>>();
                foreach (DataRow row in dt.Rows)
                {
                    var dict = new Dictionary<string, object?>();
                    foreach (DataColumn col in dt.Columns)
                    {
                        dict[col.ColumnName] = ParseValue(row[col]?.ToString() ?? "");
                    }
                    list.Add(dict);
                }
                return JsonSerializer.Serialize(list, options);
            }
        }

        // ==================== 树形结构解析 ====================

        /// <summary>
        /// 解析 JSON 为树形节点列表（用于 TreeView）
        /// </summary>
        public static List<JsonTreeNode> ParseToTree(string json)
        {
            using var doc = JsonDocument.Parse(json);
            var nodes = new List<JsonTreeNode>();
            BuildTree(doc.RootElement, "root", nodes, "");
            return nodes;
        }

        private static void BuildTree(JsonElement element, string key, List<JsonTreeNode> nodes, string path)
        {
            var currentPath = string.IsNullOrEmpty(path) ? key : $"{path}.{key}";
            var node = new JsonTreeNode
            {
                Key = key,
                Path = currentPath,
                Type = element.ValueKind.ToString()
            };

            switch (element.ValueKind)
            {
                case JsonValueKind.Object:
                    node.Value = $"{{{element.EnumerateObject().Count()} 项}}";
                    foreach (var prop in element.EnumerateObject())
                        BuildTree(prop.Value, prop.Name, node.Children, currentPath);
                    break;

                case JsonValueKind.Array:
                    node.Value = $"[{element.GetArrayLength()} 项]";
                    int idx = 0;
                    foreach (var item in element.EnumerateArray())
                    {
                        BuildTree(item, $"[{idx}]", node.Children, currentPath);
                        idx++;
                    }
                    break;

                default:
                    node.Value = GetElementValue(element);
                    break;
            }

            nodes.Add(node);
        }

        // ==================== JSON 路径提取 ====================

        /// <summary>
        /// 用简单路径表达式提取 JSON 值（如 data.items[0].name）
        /// </summary>
        public static string ExtractByPath(string json, string path)
        {
            try
            {
                using var doc = JsonDocument.Parse(json);
                var current = doc.RootElement;

                var segments = ParsePath(path);
                foreach (var seg in segments)
                {
                    if (seg.IsIndex)
                    {
                        if (current.ValueKind != JsonValueKind.Array)
                            return $"错误: '{seg.Key}' 不是数组";
                        current = current[seg.Index];
                    }
                    else
                    {
                        if (current.ValueKind != JsonValueKind.Object)
                            return $"错误: 无法在非对象类型上访问属性 '{seg.Key}'";
                        if (!current.TryGetProperty(seg.Key, out var next))
                            return $"错误: 找不到属性 '{seg.Key}'";
                        current = next;
                    }
                }

                if (current.ValueKind is JsonValueKind.Object or JsonValueKind.Array)
                {
                    return JsonSerializer.Serialize(current, new JsonSerializerOptions
                    {
                        WriteIndented = true,
                        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
                    });
                }

                return GetElementValue(current);
            }
            catch (Exception ex)
            {
                return $"提取失败: {ex.Message}";
            }
        }

        // ==================== JSON → CSV ====================

        /// <summary>
        /// 将 JSON 数组转为 CSV 字符串
        /// </summary>
        public static string JsonToCsv(string json)
        {
            var dt = ParseToDataTable(json);
            using var writer = new StringWriter();
            using var csv = new CsvWriter(writer, new CsvConfiguration(CultureInfo.InvariantCulture));

            // 写入表头
            foreach (DataColumn col in dt.Columns)
            {
                csv.WriteField(col.ColumnName);
            }
            csv.NextRecord();

            // 写入数据
            foreach (DataRow row in dt.Rows)
            {
                foreach (DataColumn col in dt.Columns)
                {
                    csv.WriteField(row[col]?.ToString() ?? "");
                }
                csv.NextRecord();
            }

            return writer.ToString();
        }

        // ==================== 辅助方法 ====================

        private static string GetElementValue(JsonElement element)
        {
            return element.ValueKind switch
            {
                JsonValueKind.String => element.GetString() ?? "",
                JsonValueKind.Number => element.GetRawText(),
                JsonValueKind.True => "true",
                JsonValueKind.False => "false",
                JsonValueKind.Null => "null",
                JsonValueKind.Object => element.GetRawText(),
                JsonValueKind.Array => element.GetRawText(),
                _ => element.GetRawText()
            };
        }

        private static object? ParseValue(string val)
        {
            if (string.IsNullOrEmpty(val) || val == "null") return null;
            if (val == "true") return true;
            if (val == "false") return false;
            if (long.TryParse(val, out long l)) return l;
            if (double.TryParse(val, out double d)) return d;
            // 尝试解析嵌套 JSON
            if ((val.StartsWith('{') && val.EndsWith('}')) || (val.StartsWith('[') && val.EndsWith(']')))
            {
                try { return JsonSerializer.Deserialize<JsonNode>(val); } catch { }
            }
            return val;
        }

        private static int GetMaxDepth(JsonElement element, int current = 0)
        {
            int max = current;
            if (element.ValueKind == JsonValueKind.Object)
            {
                foreach (var prop in element.EnumerateObject())
                    max = Math.Max(max, GetMaxDepth(prop.Value, current + 1));
            }
            else if (element.ValueKind == JsonValueKind.Array)
            {
                foreach (var item in element.EnumerateArray())
                    max = Math.Max(max, GetMaxDepth(item, current + 1));
            }
            return max;
        }

        private static int CountNodes(JsonElement element)
        {
            int count = 1;
            if (element.ValueKind == JsonValueKind.Object)
            {
                foreach (var prop in element.EnumerateObject())
                    count += CountNodes(prop.Value);
            }
            else if (element.ValueKind == JsonValueKind.Array)
            {
                foreach (var item in element.EnumerateArray())
                    count += CountNodes(item);
            }
            return count;
        }

        private static List<PathSegment> ParsePath(string path)
        {
            var segments = new List<PathSegment>();
            var parts = path.Replace("[", ".[").Split('.', StringSplitOptions.RemoveEmptyEntries);
            foreach (var part in parts)
            {
                if (part == "root") continue;
                if (part.StartsWith('[') && part.EndsWith(']'))
                {
                    var idxStr = part[1..^1];
                    if (int.TryParse(idxStr, out int idx))
                        segments.Add(new PathSegment { Key = part, IsIndex = true, Index = idx });
                }
                else
                {
                    segments.Add(new PathSegment { Key = part, IsIndex = false });
                }
            }
            return segments;
        }

        private struct PathSegment
        {
            public string Key;
            public bool IsIndex;
            public int Index;
        }
    }

    /// <summary>
    /// JSON 树形节点
    /// </summary>
    public class JsonTreeNode
    {
        public string Key { get; set; } = "";
        public string Value { get; set; } = "";
        public string Type { get; set; } = "";
        public string Path { get; set; } = "";
        public List<JsonTreeNode> Children { get; set; } = [];

        public string Display => Type is "Object" or "Array"
            ? $"{Key}: {Value}"
            : $"{Key}: {Value}  ({Type})";
    }

    /// <summary>
    /// 保持插入顺序的 HashSet
    /// </summary>
    public class LinkedHashSet<T> : IEnumerable<T> where T : notnull
    {
        private readonly Dictionary<T, int> _dict = [];
        private readonly List<T> _list = [];

        public int Count => _list.Count;

        public bool Add(T item)
        {
            if (_dict.ContainsKey(item)) return false;
            _dict[item] = _list.Count;
            _list.Add(item);
            return true;
        }

        public IEnumerator<T> GetEnumerator() => _list.GetEnumerator();
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
