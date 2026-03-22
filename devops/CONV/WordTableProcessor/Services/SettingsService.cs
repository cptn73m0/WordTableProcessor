using System.IO;
using System.Text.Json;
using WordTableProcessor.Models;

namespace WordTableProcessor.Services;

public class SettingsService
{
    private static readonly string SettingsFolder = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "WordTableProcessor");

    private static readonly string SettingsFilePath = Path.Combine(SettingsFolder, "settings.json");

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };

    private AppSettings? _cachedSettings;

    public AppSettings Load()
    {
        if (_cachedSettings != null)
        {
            return _cachedSettings;
        }

        try
        {
            if (File.Exists(SettingsFilePath))
            {
                var json = File.ReadAllText(SettingsFilePath);
                _cachedSettings = JsonSerializer.Deserialize<AppSettings>(json, JsonOptions) ?? new AppSettings();
            }
            else
            {
                _cachedSettings = new AppSettings();
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error loading settings: {ex.Message}");
            _cachedSettings = new AppSettings();
        }

        return _cachedSettings;
    }

    public void Save(AppSettings settings)
    {
        try
        {
            if (!Directory.Exists(SettingsFolder))
            {
                Directory.CreateDirectory(SettingsFolder);
            }

            var json = JsonSerializer.Serialize(settings, JsonOptions);
            File.WriteAllText(SettingsFilePath, json);
            _cachedSettings = settings;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"Error saving settings: {ex.Message}");
        }
    }

    public void SaveWindowPosition(double left, double top, double width, double height, bool isMaximized)
    {
        var settings = Load();
        settings.WindowLeft = left;
        settings.WindowTop = top;
        settings.WindowWidth = width;
        settings.WindowHeight = height;
        settings.IsMaximized = isMaximized;
        Save(settings);
    }

    public void AddRecentFile(string filePath, RecentFileType fileType)
    {
        var settings = Load();
        settings.AddRecentFile(filePath, fileType);
        Save(settings);
    }

    public void ClearRecentFiles()
    {
        var settings = Load();
        settings.ClearRecentFiles();
        Save(settings);
    }
}
