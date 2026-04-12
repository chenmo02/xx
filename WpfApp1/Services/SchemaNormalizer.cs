namespace WpfApp1.Services
{
    // Converts vendor-specific DB types into the small normalized type set that
    // the validation engine understands. This keeps rule logic DB-agnostic.
    public static class SchemaNormalizer
    {
        /// <summary>
        /// 将数据库原生类型归一化为校验引擎可识别的统一类型。
        /// 这样上层校验逻辑不需要分别处理 SQL Server 和 PostgreSQL 的命名差异。
        /// </summary>
        // Public entry used when DDL / schema query results are parsed into target columns.
        public static DvNormalizedType Normalize(string dataType, DvDbType dbType)
        {
            var t = dataType.ToLower().Trim();
            return dbType == DvDbType.SqlServer ? NormalizeSqlServer(t) : NormalizePostgreSql(t);
        }

        // SQL Server 类型映射。
        // SQL Server type normalization rules.
        private static DvNormalizedType NormalizeSqlServer(string t) => t switch
        {
            "char" or "varchar" or "nchar" or "nvarchar" or "text" or "ntext" or "xml" => DvNormalizedType.String,
            "int" or "smallint" or "tinyint" => DvNormalizedType.Integer,
            "bigint" => DvNormalizedType.Long,
            "decimal" or "numeric" or "money" or "smallmoney" or "float" or "real" => DvNormalizedType.Decimal,
            "date" => DvNormalizedType.Date,
            "datetime" or "datetime2" or "smalldatetime" or "datetimeoffset" => DvNormalizedType.DateTime,
            "time" => DvNormalizedType.Time,
            "bit" => DvNormalizedType.Boolean,
            "uniqueidentifier" => DvNormalizedType.Guid,
            _ => DvNormalizedType.Unknown
        };

        // PostgreSQL 类型映射。
        // PostgreSQL type normalization rules.
        private static DvNormalizedType NormalizePostgreSql(string t) => t switch
        {
            "character varying" or "varchar" or "character" or "char" or "text" or "citext" => DvNormalizedType.String,
            "smallint" or "int2" or "integer" or "int4" or "int" => DvNormalizedType.Integer,
            "bigint" or "int8" => DvNormalizedType.Long,
            "numeric" or "decimal" or "real" or "double precision" or "float4" or "float8" or "float" => DvNormalizedType.Decimal,
            "date" => DvNormalizedType.Date,
            "timestamp" or "timestamp without time zone" or "timestamp with time zone" or "timestamptz" => DvNormalizedType.DateTime,
            "time" or "time without time zone" or "time with time zone" or "timetz" => DvNormalizedType.Time,
            "boolean" or "bool" => DvNormalizedType.Boolean,
            "uuid" => DvNormalizedType.Guid,
            "json" or "jsonb" => DvNormalizedType.Json,
            _ => DvNormalizedType.Unknown
        };
    }
}
