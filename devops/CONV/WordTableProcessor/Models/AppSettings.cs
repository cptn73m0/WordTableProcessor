using System.Text.Json;
using System.IO;

namespace WordTableProcessor.Models;

public class AppSettings
{
    public string LastDirectory { get; set; } = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
    public double WindowWidth { get; set; } = 1000;
    public double WindowHeight { get; set; } = 700;
    public double WindowLeft { get; set; } = -1;
    public double WindowTop { get; set; } = -1;
    public bool IsMaximized { get; set; } = false;
    public List<RecentFile> RecentFiles { get; set; } = new();
    public int MaxRecentFiles { get; set; } = 10;
    public IndexesSettings IndexesSettings { get; set; } = new();
    public bool UseClassicUI { get; set; } = false;

    public void AddRecentFile(string filePath, RecentFileType fileType)
    {
        var existing = RecentFiles.FirstOrDefault(f => f.FilePath.Equals(filePath, StringComparison.OrdinalIgnoreCase));
        if (existing != null)
        {
            RecentFiles.Remove(existing);
        }

        RecentFiles.Insert(0, new RecentFile
        {
            FilePath = filePath,
            FileType = fileType,
            LastAccessed = DateTime.Now
        });

        if (RecentFiles.Count > MaxRecentFiles)
        {
            RecentFiles = RecentFiles.Take(MaxRecentFiles).ToList();
        }
    }

    public void ClearRecentFiles()
    {
        RecentFiles.Clear();
    }
}

public class RecentFile
{
    public string FilePath { get; set; } = string.Empty;
    public RecentFileType FileType { get; set; }
    public DateTime LastAccessed { get; set; }

    public string FileName => Path.GetFileName(FilePath);
    public string Directory => Path.GetDirectoryName(FilePath) ?? string.Empty;
}

public enum RecentFileType
{
    WordDocument,
    CsvFile,
    ExcelFile,
    Any
}

public class IndexesSettings
{
    public int Code { get; set; } = 1;
    public string Table1End { get; set; } = "47-02-096-01";
    public string Table2End { get; set; } = "40-03-001-12";
    public string Table3End { get; set; } = "47-02-096-01";
    public string Table4End { get; set; } = "40-03-001-12";
    public int ProcessedRows { get; set; }
    public RegionalSettings Regional { get; set; } = new();
    public IndustrySettings Industry { get; set; } = new();
}

public class RegionalSettings
{
    public int Inem { get; set; } = 5;
    public int Mat { get; set; } = 6;
}

public class IndustrySettings
{
    public int Inem { get; set; } = 8;
    public int Mat { get; set; } = 9;
}
