using System.Text.Json;
using System.IO;
using WpfApp1.Models;

namespace WpfApp1.Services;

public static class QuickNoteService
{
    private static readonly string FilePath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "CCToolbox", "quick-notes.txt");
    private static readonly string LegacyFilePath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "CCToolbox", "quick-notes.json");

    public static string Load()
    {
        try
        {
            if (File.Exists(FilePath)) return File.ReadAllText(FilePath);
            if (!File.Exists(LegacyFilePath)) return string.Empty;

            var notes = JsonSerializer.Deserialize<List<QuickNote>>(File.ReadAllText(LegacyFilePath)) ?? [];
            var content = string.Join(Environment.NewLine, notes.Select(note => note.Content));
            if (content.Length > 0) Save(content);
            return content;
        }
        catch
        {
            return string.Empty;
        }
    }

    public static void Save(string content)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(FilePath)!);
        File.WriteAllText(FilePath, content);
    }
}
