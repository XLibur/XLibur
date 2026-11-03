namespace XLibur.Operations;

/// <summary>
/// Why an operation failed. The kind tells a caller whether to fix its request, accept a gap in
/// XLibur, or report a bug; Unsupported feature and Defect are the terms of CONTEXT.md.
/// </summary>
public enum ErrorKind
{
    /// <summary>The request or the file it names is wrong: a missing sheet, a bad address, a corrupt file.</summary>
    BadInput,

    /// <summary>The path policy refused the request: a path outside every root, or a write that is off.</summary>
    Forbidden,

    /// <summary>A valid Excel construct that XLibur does not handle: an unsupported feature.</summary>
    Unsupported,

    /// <summary>A bug in XLibur or in this tool rather than in the request: a defect.</summary>
    Defect,
}

/// <param name="Code">A stable snake_case identifier, such as <c>sheet_not_found</c>.</param>
public sealed record OperationError(ErrorKind Kind, string Code, string Message)
{
    public string ToText() => $"error[{Code}] {Kind}: {Message}";
}

/// <summary>
/// The one exception an operation throws for an expected failure. Anything else that escapes an
/// operation is a defect.
/// </summary>
public sealed class OperationException(ErrorKind kind, string code, string message, Exception? inner = null)
    : Exception(message, inner)
{
    public OperationError Error { get; } = new(kind, code, message);

    public static OperationException BadInput(string code, string message, Exception? inner = null)
        => new(ErrorKind.BadInput, code, message, inner);

    public static OperationException Forbidden(string code, string message)
        => new(ErrorKind.Forbidden, code, message);

    public static OperationException Unsupported(string code, string message, Exception? inner = null)
        => new(ErrorKind.Unsupported, code, message, inner);
}
