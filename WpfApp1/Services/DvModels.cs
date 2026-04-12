using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;

namespace WpfApp1.Services
{
    /// <summary>
    /// 数据验证当前支持的数据库类型。
    /// </summary>
    public enum DvDbType { SqlServer, PostgreSql }

    /// <summary>
    /// 结构字段在进入校验引擎后会被归一化为统一类型。
    /// </summary>
    public enum DvNormalizedType
    {
        String,
        Integer,
        Long,
        Decimal,
        Date,
        DateTime,
        Time,
        Boolean,
        Guid,
        Json,
        Unknown
    }

    /// <summary>
    /// 字段映射方式：来自源字段、固定值，或直接忽略。
    /// </summary>
    public enum DvMappingType { Source, Constant, Ignore }

    /// <summary>
    /// 校验结果等级。
    /// </summary>
    public enum DvValidationLevel { Error, Warning, Info }

    /// <summary>
    /// 自动匹配字段时的命中方式。
    /// </summary>
    public enum DvMatchMethod { Exact, Normalized, PrefixStrip, Semantic, Contains, Manual }

    /// <summary>
    /// 目标字段元数据，即从 DDL 或数据库结构查询中解析出的字段定义。
    /// </summary>
    public sealed class DvTargetColumn
    {
        public int OrdinalPosition { get; init; }
        public required string ColumnName { get; init; }
        public required string OriginalDataType { get; init; }
        public DvNormalizedType NormalizedType { get; init; }
        public int? MaxLength { get; init; }       // -1 表示不限制长度
        public int? NumericPrecision { get; init; }
        public int? NumericScale { get; init; }
        public bool IsNullable { get; init; }
        public DvDbType DatabaseType { get; init; }

        public string DisplayType
        {
            get
            {
                if (MaxLength.HasValue && MaxLength.Value > 0)
                {
                    return $"{OriginalDataType}({MaxLength})";
                }

                if (NumericPrecision.HasValue && NumericScale.HasValue)
                {
                    return $"{OriginalDataType}({NumericPrecision},{NumericScale})";
                }

                if (NumericPrecision.HasValue)
                {
                    return $"{OriginalDataType}({NumericPrecision})";
                }

                return OriginalDataType;
            }
        }
    }

    /// <summary>
    /// 源数据载体，可以来自 INSERT 解析，也可以来自 Excel 导入。
    /// </summary>
    public sealed class DvSourceData
    {
        public required IReadOnlyList<string> Headers { get; init; }
        public required IReadOnlyList<IReadOnlyList<string?>> Rows { get; init; }
        public int RowCount => Rows.Count;
    }

    /// <summary>
    /// 字段映射页中的一行数据。
    /// </summary>
    public sealed class DvMappingRow : INotifyPropertyChanged
    {
        /// <summary>行号，从 1 开始。</summary>
        public int RowIndex { get; set; }

        public required string TargetColumnName { get; init; }
        public required string TargetDisplayType { get; init; }
        public bool IsRequired { get; init; }
        public bool IsUuidTarget { get; init; }

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
            set
            {
                _sourceColumnName = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(NeedsHighlight));
            }
        }

        private string? _constantValue;
        public string? ConstantValue
        {
            get => _constantValue;
            set
            {
                _constantValue = value;
                OnPropertyChanged();
            }
        }

        private DvMatchMethod _matchMethod;
        public DvMatchMethod MatchMethod
        {
            get => _matchMethod;
            set
            {
                _matchMethod = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(MatchMethodText));
            }
        }

        private int _confidence;
        public int Confidence
        {
            get => _confidence;
            set
            {
                _confidence = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(ConfidenceText));
            }
        }

        private bool _isConfirmed;
        public bool IsConfirmed
        {
            get => _isConfirmed;
            set
            {
                _isConfirmed = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(NeedsHighlight));
            }
        }

        private bool _needsConfirmation;
        public bool NeedsConfirmation
        {
            get => _needsConfirmation;
            set
            {
                _needsConfirmation = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(NeedsHighlight));
            }
        }

        private string? _matchReason;
        public string? MatchReason
        {
            get => _matchReason;
            set
            {
                _matchReason = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(MatchReasonText));
            }
        }

        /// <summary>
        /// 是否为系统自动生成字段，例如 UUID 主键。
        /// 这类字段允许默认忽略，不计入“必填未映射”。
        /// </summary>
        private bool _isAutoGenCandidate;
        public bool IsAutoGenCandidate
        {
            get => _isAutoGenCandidate;
            set
            {
                _isAutoGenCandidate = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(NeedsHighlight));
            }
        }

        private bool _wasAutoIgnored;
        public bool WasAutoIgnored
        {
            get => _wasAutoIgnored;
            set
            {
                _wasAutoIgnored = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(NeedsHighlight));
            }
        }

        public bool IsSourceMode => MappingType == DvMappingType.Source;
        public bool IsConstantMode => MappingType == DvMappingType.Constant;

        /// <summary>
        /// 提供给 DataGrid ComboBox 的字符串形式。
        /// </summary>
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
            DvMatchMethod.Semantic => "语义优先",
            DvMatchMethod.Contains => "包含（待确认）",
            DvMatchMethod.Manual => "\u624B\u52A8",
            _ => "-"
        };

        public string ConfidenceText => Confidence > 0 ? $"{Confidence}%" : "-";
        public string MatchReasonText => string.IsNullOrWhiteSpace(MatchReason) ? "-" : MatchReason!;

        /// <summary>
        /// UUID 目标字段必须由用户显式处理；其他字段仅在“必填且被忽略”时高亮。
        /// </summary>
        public bool NeedsHighlight =>
            (IsUuidTarget && !IsConfirmed) ||
            (IsRequired && !IsUuidTarget && MappingType == DvMappingType.Ignore && !IsAutoGenCandidate);

        public event PropertyChangedEventHandler? PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string? name = null)
            => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    /// <summary>
    /// 单条校验问题明细。
    /// </summary>
    public sealed class DvIssue
    {
        public int RowNumber { get; init; }

        /// <summary>主键标识，例如 "emrbaid=abc123" 或 "pk1=x, pk2=y"。</summary>
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

    /// <summary>
    /// 一次批量校验后的结果汇总。
    /// </summary>
    public sealed class DvValidationResult
    {
        public required IReadOnlyList<DvIssue> Issues { get; init; }
        public int TotalRows { get; init; }
        public int ProcessedRows { get; init; }
        public System.TimeSpan Elapsed { get; init; }
        public bool WasCancelled { get; init; }

        /// <summary>
        /// 按主键去重后的错误记录数；没有主键时按行号去重。
        /// </summary>
        public int ErrorCount => Issues
            .Where(i => i.Level == DvValidationLevel.Error)
            .Select(i => i.PrimaryKeyDisplay ?? $"row:{i.RowNumber}")
            .Distinct()
            .Count();

        /// <summary>
        /// 按主键去重后的警告记录数；没有主键时按行号去重。
        /// </summary>
        public int WarningCount => Issues
            .Where(i => i.Level == DvValidationLevel.Warning)
            .Select(i => i.PrimaryKeyDisplay ?? $"row:{i.RowNumber}")
            .Distinct()
            .Count();

        /// <summary>原始错误条数，不去重。</summary>
        public int RawErrorCount => Issues.Count(i => i.Level == DvValidationLevel.Error);

        /// <summary>原始警告条数，不去重。</summary>
        public int RawWarningCount => Issues.Count(i => i.Level == DvValidationLevel.Warning);
    }
}
