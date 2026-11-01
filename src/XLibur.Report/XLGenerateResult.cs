using System;

namespace XLibur.Report;

/// <summary>
/// The outcome of <see cref="IXLTemplate.Generate"/>.
/// </summary>
public sealed class XLGenerateResult
{
    internal XLGenerateResult(TemplateErrors parsingErrors) =>
        ParsingErrors = parsingErrors ?? throw new ArgumentNullException(nameof(parsingErrors));

    /// <summary>Everything that went wrong during generation. Empty on a clean run.</summary>
    public TemplateErrors ParsingErrors { get; }

    /// <summary>Whether any cell failed to generate.</summary>
    public bool HasErrors => ParsingErrors.Count > 0;
}
