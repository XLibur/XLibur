using System.Collections.Generic;
using XLibur.Excel;
using XLibur.Report.Expressions;

namespace XLibur.Report.Tags;

/// <summary>
/// A marker a template author writes in the options row to change how a range is generated.
/// </summary>
/// <remarks>
/// A tag acts at one of two moments, and may act at both. <see cref="TransformItems"/> runs before
/// any row is written, which is where reordering belongs; <see cref="Execute"/> runs once the rows
/// exist, which is where anything referring to the generated block belongs — a total, an
/// autofilter, a column width.
/// <para>
/// Register a tag of your own with <see cref="TagsRegister.Add{T}"/>.
/// </para>
/// </remarks>
public abstract class OptionTag
{
    /// <summary>The tag as written, including its parameters.</summary>
    public TagToken Token { get; internal set; } = new(string.Empty, new Dictionary<string, string>());

    /// <summary>
    /// The worksheet column the tag was written in. Tags that act on one column of the range —
    /// a total, a sort key — read this.
    /// </summary>
    public int Column { get; internal set; }

    /// <summary>
    /// Runs before any row is written, returning the items to generate from. The default returns
    /// them unchanged.
    /// </summary>
    public virtual IReadOnlyList<object?> TransformItems(IReadOnlyList<object?> items, ProcessingContext context) => items;

    /// <summary>
    /// Runs once the rows have been written. The default does nothing.
    /// </summary>
    public virtual void Execute(ProcessingContext context)
    {
    }
}

/// <summary>
/// What a tag is given to work with: the block that was generated, the data behind it, and the way
/// back to the expression engine.
/// </summary>
public sealed class ProcessingContext
{
    internal ProcessingContext(
        IXLWorksheet worksheet,
        IXLRange generatedRange,
        IXLRange? optionsRow,
        IReadOnlyList<object?> items,
        IExpressionEngine engine,
        ExpressionScope scope,
        TemplateErrors errors,
        IReadOnlyDictionary<int, string> columnExpressions)
    {
        Worksheet = worksheet;
        GeneratedRange = generatedRange;
        OptionsRow = optionsRow;
        Items = items;
        Engine = engine;
        Scope = scope;
        Errors = errors;
        ColumnExpressions = columnExpressions;
    }

    /// <summary>The sheet being generated.</summary>
    public IXLWorksheet Worksheet { get; }

    /// <summary>The rows the range produced, excluding the options row.</summary>
    public IXLRange GeneratedRange { get; }

    /// <summary>The options row, or <c>null</c> once it has been removed.</summary>
    public IXLRange? OptionsRow { get; }

    /// <summary>The data the range was generated from, in the order it was written.</summary>
    public IReadOnlyList<object?> Items { get; }

    /// <summary>The engine evaluating this template.</summary>
    public IExpressionEngine Engine { get; }

    /// <summary>The scope holding the workbook-wide variables and this range's collection.</summary>
    public ExpressionScope Scope { get; }

    /// <summary>Where a tag reports a problem instead of throwing.</summary>
    public TemplateErrors Errors { get; }

    /// <summary>
    /// The template expression each column held, keyed by worksheet column. A column-placed tag
    /// uses it to work out what that column means — which is how <c>&lt;&lt;Sort&gt;&gt;</c> knows
    /// what to sort by without being told twice.
    /// </summary>
    public IReadOnlyDictionary<int, string> ColumnExpressions { get; }
}
