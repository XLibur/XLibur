using System.Text.Json;

namespace PerfReport;

/// <summary>A record on disk has a schema this tool does not know, so it cannot be compared safely.</summary>
public sealed class UnsupportedSchemaException(string file, int schema)
    : Exception($"{file} has schema {schema}; this tool only reads schema {ReleaseRecord.CurrentSchema}.");

/// <summary>The <c>records/*.json</c> files of the perf-history branch.</summary>
public static class HistoryStore
{
    public static readonly JsonSerializerOptions Json = new(JsonSerializerDefaults.Web) { WriteIndented = true };

    /// <summary>Loads every record, lowest version first (see <see cref="VersionOrder"/>). Refuses a record with an unknown schema.</summary>
    public static List<ReleaseRecord> Load(string historyDir)
    {
        var dir = Path.Combine(historyDir, "records");
        if (!Directory.Exists(dir))
            return [];

        var records = new List<ReleaseRecord>();
        foreach (var file in Directory.GetFiles(dir, "*.json"))
        {
            var text = File.ReadAllText(file);
            using (var doc = JsonDocument.Parse(text))
            {
                var schema = doc.RootElement.TryGetProperty("schema", out var s) && s.TryGetInt32(out var n) ? n : 0;
                if (schema != ReleaseRecord.CurrentSchema)
                    throw new UnsupportedSchemaException(Path.GetFileName(file), schema);
            }

            records.Add(JsonSerializer.Deserialize<ReleaseRecord>(text, Json)
                        ?? throw new InvalidDataException($"{file} is empty."));
        }

        return VersionOrder.Sort(records);
    }

    /// <summary>Writes the record. A second run for the same version replaces the first.</summary>
    public static string Save(string historyDir, ReleaseRecord record)
    {
        var dir = Path.Combine(historyDir, "records");
        Directory.CreateDirectory(dir);
        var name = record.Manual
            ? "manual-" + new string(record.RecordedAt.Where(char.IsAsciiDigit).ToArray())
            : Sanitize(record.Version);
        var path = Path.Combine(dir, name + ".json");
        File.WriteAllText(path, JsonSerializer.Serialize(record, Json) + "\n");
        return path;
    }

    private static string Sanitize(string version) =>
        new(version.Select(c => char.IsAsciiLetterOrDigit(c) || c is '.' or '-' or '+' ? c : '_').ToArray());
}
