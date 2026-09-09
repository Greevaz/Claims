using System.IO;
using System.Text.Json;

namespace ClaimsCalculatorApp;

public sealed class RateChangesStore
{
    private readonly string _filePath = Path.Combine(
        AppContext.BaseDirectory,
        "rate-changes.json");

    public List<RateChange> Load(string az)
    {
        if (string.IsNullOrWhiteSpace(az) || !File.Exists(_filePath))
            return new List<RateChange>();

        var entries = JsonSerializer.Deserialize<List<RateChangeSettings>>(
            File.ReadAllText(_filePath)) ?? new List<RateChangeSettings>();
        var match = entries.FirstOrDefault(entry =>
            string.Equals(entry.Az, az.Trim(), StringComparison.OrdinalIgnoreCase));

        return match?.Changes.Select(change => new RateChange
        {
            EffectiveDate = change.EffectiveDate,
            NewRatePercent = change.NewRatePercent
        }).ToList() ?? new List<RateChange>();
    }

    public void Save(string az, IEnumerable<RateChange> changes)
    {
        if (string.IsNullOrWhiteSpace(az))
            return;

        var directory = Path.GetDirectoryName(_filePath)
            ?? throw new InvalidOperationException("Could not determine the settings directory.");
        Directory.CreateDirectory(directory);

        var entries = File.Exists(_filePath)
            ? JsonSerializer.Deserialize<List<RateChangeSettings>>(File.ReadAllText(_filePath))
                ?? new List<RateChangeSettings>()
            : new List<RateChangeSettings>();
        var entry = entries.FirstOrDefault(item =>
            string.Equals(item.Az, az.Trim(), StringComparison.OrdinalIgnoreCase));

        if (entry is null)
        {
            entry = new RateChangeSettings { Az = az.Trim() };
            entries.Add(entry);
        }

        entry.Changes = changes.Select(change => new RateChangeSetting
        {
            EffectiveDate = change.EffectiveDate,
            NewRatePercent = change.NewRatePercent
        }).ToList();

        var json = JsonSerializer.Serialize(entries, new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(_filePath, json);
    }

    private sealed class RateChangeSettings
    {
        public string Az { get; set; } = "";
        public List<RateChangeSetting> Changes { get; set; } = new();
    }

    private sealed class RateChangeSetting
    {
        public DateTime EffectiveDate { get; set; }
        public decimal NewRatePercent { get; set; }
    }
}
