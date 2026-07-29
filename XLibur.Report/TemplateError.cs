using System;

namespace XLibur.Report;

/// <summary>
/// One problem found while generating a report — a malformed expression, an expression that threw,
/// or a tag that could not be applied.
/// </summary>
/// <remarks>
/// Generation records errors rather than throwing, so a single bad cell costs that cell and not
/// the whole report. The cell itself is left showing the error text.
/// </remarks>
public sealed class TemplateError
{
    /// <summary>Creates an error describing a failure at <paramref name="cellAddress"/>.</summary>
    public TemplateError(string message, string? sheetName = null, string? cellAddress = null, Exception? exception = null)
    {
        Message = message ?? throw new ArgumentNullException(nameof(message));
        SheetName = sheetName;
        CellAddress = cellAddress;
        Exception = exception;
    }

    /// <summary>What went wrong.</summary>
    public string Message { get; }

    /// <summary>The worksheet the failure occurred on, when known.</summary>
    public string? SheetName { get; }

    /// <summary>The address of the offending cell (for example <c>B7</c>), when known.</summary>
    public string? CellAddress { get; }

    /// <summary>The underlying exception, when the failure came from one.</summary>
    public Exception? Exception { get; }

    /// <summary>The location as a sheet-qualified address, or an empty string when unknown.</summary>
    public string Location => (SheetName, CellAddress) switch
    {
        (null, null) => string.Empty,
        (null, { } cell) => cell,
        ({ } sheet, null) => sheet,
        ({ } sheet, { } cell) => $"{sheet}!{cell}",
    };

    /// <inheritdoc />
    public override string ToString() => Location.Length == 0 ? Message : $"{Location}: {Message}";
}
