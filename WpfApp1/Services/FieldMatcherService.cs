using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace WpfApp1.Services
{
    public static class FieldMatcherService
    {
        private static readonly string[] DefaultPrefixes = ["zjb_", "tmp_", "mid_", "src_"];
        private static readonly string[] DefaultSuffixes = ["_text", "_str", "_value"];

        /// <summary>自动生成字段映射建议。</summary>
        public static List<DvMappingRow> AutoMap(
            IReadOnlyList<DvTargetColumn> targets,
            IReadOnlyList<string> sourceHeaders)
        {
            // 追踪已占用的源字段（避免一对多）
            var usedSources = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var result = new List<DvMappingRow>();
            int rowIdx = 0;

            foreach (var target in targets)
            {
                var (srcName, method, conf) = BestMatch(target.ColumnName, sourceHeaders, usedSources);

                rowIdx++;
                var row = new DvMappingRow
                {
                    RowIndex = rowIdx,
                    TargetColumnName = target.ColumnName,
                    TargetDisplayType = target.DisplayType,
                    IsRequired = !target.IsNullable
                };

                // 判断是否为系统自动生成字段（uuid/uniqueidentifier 类型）
                bool isAutoGen = target.NormalizedType == DvNormalizedType.Guid;

                if (srcName != null)
                {
                    row.MappingType = DvMappingType.Source;
                    row.SourceColumnName = srcName;
                    row.MatchMethod = method;
                    row.Confidence = conf;
                    row.IsConfirmed = method != DvMatchMethod.Contains; // Contains 需手动确认
                    row.NeedsConfirmation = method == DvMatchMethod.Contains;
                    row.IsAutoGenCandidate = isAutoGen;
                    usedSources.Add(srcName);
                }
                else
                {
                    row.MappingType = DvMappingType.Ignore;
                    row.MatchMethod = DvMatchMethod.Manual;
                    row.Confidence = 0;
                    row.IsAutoGenCandidate = isAutoGen;
                    // 自动生成字段（uuid）忽略时自动确认；普通可空字段忽略也确认
                    row.IsConfirmed = isAutoGen || target.IsNullable;
                }

                result.Add(row);
            }

            return result;
        }

        private static (string? source, DvMatchMethod method, int conf) BestMatch(
            string target, IReadOnlyList<string> sources, HashSet<string> used)
        {
            var available = sources.Where(s => !used.Contains(s)).ToList();
            if (available.Count == 0) return (null, DvMatchMethod.Manual, 0);

            string tLower = target.ToLowerInvariant();
            string tNorm = Normalize(tLower);

            // Level 1: 精确匹配
            foreach (var s in available)
                if (string.Equals(s, target, StringComparison.OrdinalIgnoreCase))
                    return (s, DvMatchMethod.Exact, 100);

            // Level 2: 标准化匹配
            foreach (var s in available)
                if (Normalize(s) == tNorm)
                    return (s, DvMatchMethod.Normalized, 90);

            // Level 3: 前后缀剥离
            string tStripped = Strip(tLower);
            foreach (var s in available)
                if (Strip(Normalize(s)) == tStripped)
                    return (s, DvMatchMethod.PrefixStrip, 75);

            // Level 4: 包含匹配（仅候选，confidence 低）
            foreach (var s in available)
            {
                var sn = Normalize(s.ToLowerInvariant());
                if (sn.Contains(tNorm) || tNorm.Contains(sn))
                    return (s, DvMatchMethod.Contains, 50);
            }

            return (null, DvMatchMethod.Manual, 0);
        }

        private static string Normalize(string s)
            => Regex.Replace(s.ToLowerInvariant(), @"[\-_\s]", "");

        private static string Strip(string s)
        {
            foreach (var p in DefaultPrefixes)
                if (s.StartsWith(p)) { s = s[p.Length..]; break; }
            foreach (var sf in DefaultSuffixes)
                if (s.EndsWith(sf)) { s = s[..^sf.Length]; break; }
            return s;
        }
    }
}
