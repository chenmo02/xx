namespace WpfApp1.Models;

public sealed class QuickNote
{
    public string Content { get; set; } = string.Empty;
    public DateTime CreatedAt { get; set; } = DateTime.Now;
}
