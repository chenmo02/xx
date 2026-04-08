using System.IO;
using System.Text;
using System.Text.Json;

namespace WpfApp1.Services
{
    public sealed class ImportSettings
    {
        public string DefaultDbType { get; set; } = "PostgreSQL";

        public string DefaultTableName { get; set; } = "TempTable";

        public string? TempTablePrefix { get; set; }

        public int BatchSize { get; set; } = 1000;

        public bool DropIfExists { get; set; } = true;

        public bool BatchInsert { get; set; } = true;

        public bool LimitFieldLength { get; set; } = true;

        public string DefaultExportPath { get; set; } = string.Empty;
    }

    public static class ImportSettingsService
    {
        private static readonly JsonSerializerOptions SerializerOptions = new()
        {
            WriteIndented = true
        };

        public static event EventHandler<ImportSettings>? SettingsSaved;

        public static string ConfigPath => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "settings.json");

        public static ImportSettings Load()
        {
            try
            {
                if (!File.Exists(ConfigPath))
                {
                    return new ImportSettings();
                }

                string json = File.ReadAllText(ConfigPath, Encoding.UTF8);
                var settings = JsonSerializer.Deserialize<ImportSettings>(json, SerializerOptions) ?? new ImportSettings();
                return Normalize(settings);
            }
            catch
            {
                return new ImportSettings();
            }
        }

        public static void Save(ImportSettings settings)
        {
            var normalized = Normalize(Clone(settings));

            string? directory = Path.GetDirectoryName(ConfigPath);
            if (!string.IsNullOrWhiteSpace(directory))
            {
                Directory.CreateDirectory(directory);
            }

            File.WriteAllText(ConfigPath, JsonSerializer.Serialize(normalized, SerializerOptions), Encoding.UTF8);
            SettingsSaved?.Invoke(null, Clone(normalized));
        }

        public static ImportSettings Normalize(ImportSettings settings)
        {
            settings.DefaultDbType = NormalizeDbType(settings.DefaultDbType);
            settings.DefaultTableName = NormalizeTableName(settings.DefaultDbType, settings.DefaultTableName, settings.TempTablePrefix);
            settings.BatchSize = settings.BatchSize <= 0 ? 1000 : settings.BatchSize;
            settings.DefaultExportPath = settings.DefaultExportPath?.Trim() ?? string.Empty;
            return settings;
        }

        private static string NormalizeDbType(string? dbType)
        {
            string normalized = dbType?.Trim() ?? string.Empty;
            return normalized switch
            {
                "SQLServer" => "SQL Server",
                "SQL Server" => "SQL Server",
                "MySQL" => "MySQL",
                "Oracle" => "Oracle",
                _ => "PostgreSQL"
            };
        }

        private static ImportSettings Clone(ImportSettings settings) => new()
        {
            DefaultDbType = settings.DefaultDbType,
            DefaultTableName = settings.DefaultTableName,
            TempTablePrefix = settings.TempTablePrefix,
            BatchSize = settings.BatchSize,
            DropIfExists = settings.DropIfExists,
            BatchInsert = settings.BatchInsert,
            LimitFieldLength = settings.LimitFieldLength,
            DefaultExportPath = settings.DefaultExportPath
        };

        private static string NormalizeTableName(string normalizedDbType, string? defaultTableName, string? legacyPrefix)
        {
            SqlGeneratorService.DbType dbType = normalizedDbType switch
            {
                "SQL Server" => SqlGeneratorService.DbType.SqlServer,
                "MySQL" => SqlGeneratorService.DbType.MySQL,
                "Oracle" => SqlGeneratorService.DbType.Oracle,
                _ => SqlGeneratorService.DbType.PostgreSQL
            };

            string candidate = string.IsNullOrWhiteSpace(defaultTableName)
                ? legacyPrefix?.Trim() ?? string.Empty
                : defaultTableName.Trim();

            return SqlGeneratorService.NormalizeTableName(dbType, candidate, legacyPrefix);
        }
    }
}
