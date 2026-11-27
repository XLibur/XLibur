using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace XLibur.Excel.IO;

/// <summary>
/// The rich value of an error, [MS-XLSX] 2.3.6.1.3: the keys of its <c>_error</c> structure, and a
/// value for each. Excel 365 writes keys beyond <c>errorType</c>, such as <c>subType</c> and
/// <c>argument</c>, for the error of a LAMBDA, and a save writes them as loaded while the cell holds
/// the error (#954). Two are equal when their error, keys and values are.
/// </summary>
internal sealed class ErrorRichValue : IEquatable<ErrorRichValue>
{
    internal ErrorRichValue(XLError error, IReadOnlyList<ErrorRichValueKey> keys, IReadOnlyList<string> values)
    {
        Error = error;
        Keys = keys;
        Values = values;
    }

    /// <summary>The error its <c>errorType</c> names.</summary>
    internal XLError Error { get; }

    /// <summary>The keys of its structure, in order.</summary>
    internal IReadOnlyList<ErrorRichValueKey> Keys { get; }

    /// <summary>The values, in the order of the keys.</summary>
    internal IReadOnlyList<string> Values { get; }

    /// <summary>The rich value XLibur writes for an error that has none: its <c>errorType</c> alone.</summary>
    internal static ErrorRichValue Of(XLError error) =>
        new(error, [new ErrorRichValueKey(RichDataWriter.ErrorTypeKey, "i")],
            [((int)error).ToString(CultureInfo.InvariantCulture)]);

    public bool Equals(ErrorRichValue? other) =>
        other is not null && Error == other.Error && Keys.SequenceEqual(other.Keys) && Values.SequenceEqual(other.Values);

    public override bool Equals(object? obj) => Equals(obj as ErrorRichValue);

    public override int GetHashCode() => HashCode.Combine(Error, Keys.Count, Values.Count > 0 ? Values[^1] : null);
}

/// <summary>A key of an <c>_error</c> structure: its name, and its type, such as <c>i</c> or <c>s</c>.</summary>
internal readonly record struct ErrorRichValueKey(string Name, string? Type);
