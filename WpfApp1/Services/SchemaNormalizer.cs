namespace WpfApp1.Services
{
    public static class SchemaNormalizer
    {
        public static DvNormalizedType Normalize(string dataType, DvDbType dbType)
        {
            var t = dataType.ToLower().Trim();
            return dbType == DvDbType.SqlServer ? NormalizeSqlServer(t) : NormalizePostgreSql(t);
        }

        private static DvNormalizedType NormalizeSqlServer(string t) => t switch
        {
            "char" or "varchar" or "nchar" or "nvarchar" or "text" or "ntext" or "xml" => DvNormalizedType.String,
            "int" or "smallint" or "tinyint" => DvNormalizedType.Integer,
            "bigint" => DvNormalizedType.Long,
            "decimal" or "numeric" or "money" or "smallmoney" or "float" or "real" => DvNormalizedType.Decimal,
            "date" => DvNormalizedType.Date,
            "datetime" or "datetime2" or "smalldatetime" or "datetimeoffset" => DvNormalizedType.DateTime,
            "bit" => DvNormalizedType.Boolean,
            "uniqueidentifier" => DvNormalizedType.Guid,
            _ => DvNormalizedType.Unknown
        };

        private static DvNormalizedType NormalizePostgreSql(string t) => t switch
        {
            "character varying" or "varchar" or "character" or "char" or "text" or "citext" => DvNormalizedType.String,
            "smallint" or "int2" or "integer" or "int4" or "int" => DvNormalizedType.Integer,
            "bigint" or "int8" => DvNormalizedType.Long,
            "numeric" or "decimal" or "real" or "double precision" or "float4" or "float8" or "float" => DvNormalizedType.Decimal,
            "date" => DvNormalizedType.Date,
            "timestamp" or "timestamp without time zone" or "timestamp with time zone" or "timestamptz" => DvNormalizedType.DateTime,
            "boolean" or "bool" => DvNormalizedType.Boolean,
            "uuid" => DvNormalizedType.Guid,
            "json" or "jsonb" => DvNormalizedType.Json,
            _ => DvNormalizedType.Unknown
        };
    }
}
