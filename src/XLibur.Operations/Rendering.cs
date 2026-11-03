using System.Globalization;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;
using XLibur.Excel;

namespace XLibur.Operations;

public static class Rendering
{
    /// <summary>
    /// The JSON shape of every request and result: camelCase, enums as camelCase strings, nulls left
    /// out, and a request property the type does not have is an error rather than silently ignored.
    /// </summary>
    public static JsonSerializerOptions JsonOptions { get; } = CreateOptions();

    private static JsonSerializerOptions CreateOptions()
    {
        var options = new JsonSerializerOptions
        {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
            UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow,
            Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
            RespectNullableAnnotations = true,
        };
        options.Converters.Add(new JsonStringEnumConverter(JsonNamingPolicy.CamelCase));
        options.Converters.Add(new CellValueJsonConverter());
        options.MakeReadOnly(populateMissingResolver: true);
        return options;
    }

    internal static string RenderResult(IOperationResult result, ResultFormat format)
    {
        if (format == ResultFormat.Json)
            return JsonSerializer.Serialize(result, result.GetType(), JsonOptions);

        var text = new StringBuilder();
        result.WriteText(text);
        return text.ToString();
    }

    internal static string RenderError(OperationError error, ResultFormat format) => format == ResultFormat.Json
        ? JsonSerializer.Serialize(new { error }, JsonOptions)
        : error.ToText();

    /// <summary>
    /// A value as one TSV field: numbers in invariant shortest round-trip form, dates in ISO 8601,
    /// errors as Excel shows them, and tab, newline, return and backslash escaped in text.
    /// </summary>
    public static string ToText(XLCellValue value) => value.Type switch
    {
        XLDataType.Blank => "",
        XLDataType.Boolean => value.GetBoolean() ? "TRUE" : "FALSE",
        XLDataType.Number => value.GetNumber().ToString(CultureInfo.InvariantCulture),
        XLDataType.Text => Escape(value.GetText()),
        XLDataType.Error => ErrorText(value.GetError()),
        XLDataType.DateTime => FormatDate(value.GetDateTime()),
        XLDataType.TimeSpan => value.GetTimeSpan().ToString("c", CultureInfo.InvariantCulture),
        _ => throw new ArgumentOutOfRangeException(nameof(value)),
    };

    /// <summary>An enum member as JSON writes it, such as <c>errorValue</c>, so the text form uses the same word.</summary>
    public static string Name<T>(T value) where T : struct, Enum => JsonNamingPolicy.CamelCase.ConvertName(value.ToString());

    /// <summary>An error value as Excel shows it, such as <c>#DIV/0!</c>.</summary>
    public static string ErrorText(XLError error) => ((XLCellValue)error).ToString(CultureInfo.InvariantCulture);

    public static string Escape(string text)
    {
        if (text.AsSpan().IndexOfAny("\t\n\r\\") < 0)
            return text;

        return new StringBuilder(text)
            .Replace(@"\", @"\\")
            .Replace("\t", @"\t")
            .Replace("\n", @"\n")
            .Replace("\r", @"\r")
            .ToString();
    }

    internal static string FormatDate(DateTime date) => date.TimeOfDay == TimeSpan.Zero
        ? date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
        : date.ToString("yyyy-MM-dd'T'HH:mm:ss.FFF", CultureInfo.InvariantCulture);

    /// <summary>A cell's sheet-qualified A1 address, with the sheet name quoted when Excel would quote it.</summary>
    public static string Address(IXLCell cell) => cell.Address.ToStringRelative(includeSheet: true);

    /// <summary>A range's A1 address without its sheet: <c>B2:D9</c>, or <c>B2</c> for one cell.</summary>
    public static string Area(IXLRangeAddress range)
        => range.FirstAddress.Equals(range.LastAddress)
            ? range.FirstAddress.ToStringRelative()
            : range.ToStringRelative();
}

/// <summary>
/// A cell value in JSON: blank is null, a boolean, number or text is itself, and the types JSON
/// lacks are tagged objects: <c>{"error":"#N/A"}</c>, <c>{"date":"2026-01-31"}</c>, <c>{"time":"01:30:00"}</c>.
/// </summary>
internal sealed class CellValueJsonConverter : JsonConverter<XLCellValue>
{
    public override XLCellValue Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
    {
        switch (reader.TokenType)
        {
            case JsonTokenType.Null:
                return Blank.Value;
            case JsonTokenType.True:
                return true;
            case JsonTokenType.False:
                return false;
            case JsonTokenType.Number:
                return reader.GetDouble();
            case JsonTokenType.String:
                return reader.GetString();
            case JsonTokenType.StartObject:
                using (var document = JsonDocument.ParseValue(ref reader))
                    return FromTagged(document.RootElement);
            default:
                throw new JsonException($"A cell value cannot be a JSON {reader.TokenType}.");
        }
    }

    private static XLCellValue FromTagged(JsonElement element)
    {
        if (element.TryGetProperty("error", out var error) && XLErrorParsing.TryParse(error.GetString(), out var parsed))
            return parsed;
        if (element.TryGetProperty("date", out var date) &&
            DateTime.TryParse(date.GetString(), CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var dateTime))
            return dateTime;
        if (element.TryGetProperty("time", out var time) &&
            TimeSpan.TryParse(time.GetString(), CultureInfo.InvariantCulture, out var timeSpan))
            return timeSpan;
        throw new JsonException("A tagged cell value is {\"error\":\"#N/A\"}, {\"date\":\"2026-01-31\"} or {\"time\":\"01:30:00\"}.");
    }

    public override void Write(Utf8JsonWriter writer, XLCellValue value, JsonSerializerOptions options)
    {
        switch (value.Type)
        {
            case XLDataType.Blank:
                writer.WriteNullValue();
                break;
            case XLDataType.Boolean:
                writer.WriteBooleanValue(value.GetBoolean());
                break;
            case XLDataType.Number:
                writer.WriteNumberValue(value.GetNumber());
                break;
            case XLDataType.Text:
                writer.WriteStringValue(value.GetText());
                break;
            case XLDataType.Error:
                writer.WriteStartObject();
                writer.WriteString("error", Rendering.ErrorText(value.GetError()));
                writer.WriteEndObject();
                break;
            case XLDataType.DateTime:
                writer.WriteStartObject();
                writer.WriteString("date", Rendering.FormatDate(value.GetDateTime()));
                writer.WriteEndObject();
                break;
            case XLDataType.TimeSpan:
                writer.WriteStartObject();
                writer.WriteString("time", value.GetTimeSpan().ToString("c", CultureInfo.InvariantCulture));
                writer.WriteEndObject();
                break;
        }
    }
}

internal static class XLErrorParsing
{
    private static readonly Dictionary<string, XLError> ByText = Enum.GetValues<XLError>()
        .GroupBy(Rendering.ErrorText, StringComparer.OrdinalIgnoreCase)
        .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);

    internal static bool TryParse(string? text, out XLError error)
    {
        error = default;
        return text is not null && ByText.TryGetValue(text, out error);
    }
}
