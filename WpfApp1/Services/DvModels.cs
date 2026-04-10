using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;

namespace WpfApp1.Services
{
    public enum DvDbType { SqlServer, PostgreSql }

    public enum DvNormalizedType
    {
        String, Integer, Long, Decimal, Date, DateTime, Boolean, Guid, Json, Unknown
    }

    public enum DvMappingType { Source, Constant, Ignore }

    public enum DvValidationLevel { Error, Warning, Info }

    public enum DvMatchMethod { Exact, Normalized, PrefixStrip, Contains, Manual }

    // ─── 目标字段元数据 ──────────────────────────────────────
    public sealed class DvTargetColumn
    {
        public int OrdinalPosition { get; init; }
        public required string ColumnName { get; init; }
        public required string OriginalDataType { get; init; }
        public DvNormalizedType NormalizedType { get; init; }
        public int? MaxLength { get; init; }       // -1 = unlimited
        public int? NumericPrecision { get; init; }
        public int? NumericScale { get; init; }
        public bool IsNullable { get; init; }
        public DvDbType DatabaseType { get; init; }

        public string DisplayType
        {
            get
            {
                if (MaxLength.HasValue && MaxLength.Value > 0)
                    return $"{OriginalDataType}({MaxLength})";
                if (NumericPrecision.HasValue && NumericScale.HasValue)
                    return $"{OriginalDataType}({NumericPrecision},{NumericScale})";
                if (NumericPrecision.HasValue)
                    return $"{OriginalDataType}({NumericPrecision})";
                return OriginalDataType;
            }
        }
    }

    // ─── 数据（来自 INSERT 或 Excel）──────────────────────────
    public sealed class DvSourceData
    {
        public required IReadOnlyList<string> Headers { get; init; }
        public required IReadOnlyList<IReadOnlyList<string?>> Rows { get; init; }
        public int RowCount => Rows.Count;
    }

    // ─── 字段映射行（DataGrid 绑定）─────────────────────────
    public sealed class DvMappingRow : INotifyPropertyChanged
    {
        /// <summary>行号（从 1 开始）</summary>
        public int RowIndex { get; set; }

        public required string TargetColumnName { get; init; }
        public required string TargetDisplayType { get; init; }
        public bool IsRequired { get; init; }

        private DvMappingType _mappingType = DvMappingType.Ignore;
        public DvMappingType MappingType
        {
            get => _mappingType;
            set
        {
            _mappingType = value;
            OnPropertyChanged();
            OnPropertyChanged(nameof(IsSourceMode));
            OnPropertyChanged(nameof(IsConstantMode));
            OnPropertyChanged(nameof(MappingTypeStr));
            OnPropertyChanged(nameof(NeedsHighlight));
        }
        }

        private string? _sourceColumnName;
        public string? SourceColumnName
        {
            get => _sourceColumnName;
            set { _sourceColumnName = value; OnPropertyChanged(); }
        }

        private string? _constantValue;
        public string? ConstantValue
        {
            get => _constantValue;
            set { _constantValue = value; OnPropertyChanged(); }
        }

        public DvMatchMethod MatchMethod { get; set; }
        public int Confidence { get; set; }

        private bool _isConfirmed;
        public bool IsConfirmed
        {
            get => _isConfirmed;
            set { _isConfirmed = value; OnPropertyChanged(); }
        }

        public bool NeedsConfirmation { get; set; }

        /// <summary>是否为系统自动生成字段（如 uuid 主键），可安全忽略</summary>
        public bool IsAutoGenCandidate { get; set; }

        public bool IsSourceMode => MappingType == DvMappingType.Source;
        public bool IsConstantMode => MappingType == DvMappingType.Constant;

        /// <summary>映射方式字符串（供 DataGrid ComboBox 双向绑定）</summary>
        public string MappingTypeStr
        {
            get => MappingType switch
            {
                DvMappingType.Source => "源字段映射",
                DvMappingType.Constant => "固定值",
                _ => "忽略"
            };
            set
            {
                MappingType = value switch
                {
                    "源字段映射" => DvMappingType.Source,
                    "固定值" => DvMappingType.Constant,
                    _ => DvMappingType.Ignore
                };
            }
        }

        public string MatchMethodText => MatchMethod switch
        {
            DvMatchMethod.Exact => "精确",
            DvMatchMethod.Normalized => "标准化",
            DvMatchMethod.PrefixStrip => "前缀剥离",
            DvMatchMethod.Contains => "包含(待确认)",
            DvMatchMethod.Manual => "-",
            _ => "-"
        };

        public string ConfidenceText => Confidence > 0 ? $"{Confidence}%" : "-";

        /// <summary>必填且为忽略模式 → 高亮提醒（自动生成字段除外）</summary>
        public bool NeedsHighlight => IsRequired && MappingType == DvMappingType.Ignore && !IsAutoGenCandidate;

        public event PropertyChangedEventHandler? PropertyChanged;
        private void OnPropertyChanged([CallerMemberName] string? name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    // ─── 校验问题 ────────────────────────────────────────────
    public sealed class DvIssue
    {
        public int RowNumber { get; init; }
        /// <summary>主键标识，如 "emrbaid=abc123" 或 "pk1=x, pk2=y"</summary>
        public string? PrimaryKeyDisplay { get; set; }
        public string? SourceColumnName { get; init; }
        public required string TargetColumnName { get; init; }
        public required string TargetDataType { get; init; }
        public DvValidationLevel Level { get; init; }
        public required string ErrorType { get; init; }
        public string? ActualValue { get; init; }
        public required string Message { get; init; }

        public string LevelText => Level switch
        {
            DvValidationLevel.Error => "错误",
            DvValidationLevel.Warning => "警告",
            DvValidationLevel.Info => "提示",
            _ => ""
        };
    }

    // ─── 校验结果 ────────────────────────────────────────────
    public sealed class DvValidationResult
    {
        public required IReadOnlyList<DvIssue> Issues { get; init; }
        public int TotalRows { get; init; }
        public int ProcessedRows { get; init; }
        public System.TimeSpan Elapsed { get; init; }
        public bool WasCancelled { get; init; }

        /// <summary>按主键去重的错误记录数（无主键时按行号去重）</summary>
        public int ErrorCount => Issues
            .Where(i => i.Level == DvValidationLevel.Error)
            .Select(i => i.PrimaryKeyDisplay ?? $"row:{i.RowNumber}")
            .Distinct()
            .Count();

        /// <summary>按主键去重的警告记录数（无主键时按行号去重）</summary>
        public int WarningCount => Issues
            .Where(i => i.Level == DvValidationLevel.Warning)
            .Select(i => i.PrimaryKeyDisplay ?? $"row:{i.RowNumber}")
            .Distinct()
            .Count();

        /// <summary>原始错误条数（不去重）</summary>
        public int RawErrorCount => Issues.Count(i => i.Level == DvValidationLevel.Error);
        /// <summary>原始警告条数（不去重）</summary>
        public int RawWarningCount => Issues.Count(i => i.Level == DvValidationLevel.Warning);
    }
}
