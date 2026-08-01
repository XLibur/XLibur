using System.Collections.Generic;
using XLibur.Excel;
using XLibur.Report.Expressions;

namespace XLibur.Report.Tags;

/// <summary>
/// A marker a template author writes in a range's options slot to change how it is generated.
/// </summary>
/// <remarks>
/// <para>
/// A tag acts at one of two moments, and may act at both. <see cref="TransformItems"/> runs before
/// any row is written, which is where reordering and filtering belong; <see cref="Execute"/> runs
/// once the rows exist, which is where anything referring to the generated block belongs — a total,
/// an autofilter, a column width.
/// </para>
/// <para>
/// Register a tag of your own with <see cref="TagsRegister.Add{T}"/>.
/// </para>
/// </remarks>
public abstract class OptionTag
{
    /// <summary>The tag as written, including its parameters.</summary>
    public TagToken Token { get; internal set; } = new(string.Empty, new Dictionary<string, string>());

    /// <summary>The worksheet row the tag was written in.</summary>
    public int Row { get; internal set; }

    /// <summary>The worksheet column the tag was written in.</summary>
    public int Column { get; internal set; }

    /// <summary>
    /// The line of the range the tag sits in — its <see cref="Column"/> for a range that repeats
    /// downwards, its <see cref="Row"/> for one that repeats across.
    /// </summary>
    /// <remarks>
    /// A tag acting on one line of the range — a total, a sort key, a hidden column — reads this
    /// rather than <see cref="Column"/>, and works whichever way the range runs.
    /// </remarks>
    public int Line { get; internal set; }

    /// <summary>
    /// Whether the tag was written in one of the range's repeated slots rather than in its options
    /// slot.
    /// </summary>
    /// <remarks>
    /// Most tags describe the range and belong in the options slot. A tag in a repeated slot is
    /// describing <em>one item</em>, and a tag that means something at both scales reads this to tell
    /// which was meant — <c>&lt;&lt;If&gt;&gt;</c> in a repeated slot drops that item, and in the
    /// options slot drops the range.
    /// </remarks>
    public bool InRepeatedRow { get; internal set; }

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
#pragma warning disable S107 // One field per piece of processing state; a parameter object here is the same list behind a name
    internal ProcessingContext(
        IXLWorksheet worksheet,
        IXLRange generatedRange,
        IXLRange? optionsRow,
        IReadOnlyList<object?> items,
        IExpressionEngine engine,
        ExpressionScope scope,
        TemplateErrors errors,
        IReadOnlyDictionary<int, string> lineExpressions,
        Ranges.RangeAxis axis,
        IReadOnlyList<OptionTag> tags,
        List<Rewriting.PivotRequest> pendingPivots,
        bool grandTotalsDisabled = false)
#pragma warning restore S107
    {
        Tags = tags;
        PendingPivots = pendingPivots;
        Worksheet = worksheet;
        GeneratedRange = generatedRange;
        OptionsRow = optionsRow;
        Items = items;
        Engine = engine;
        Scope = scope;
        Errors = errors;
        LineExpressions = lineExpressions;
        Axis = axis;
        GrandTotalsDisabled = grandTotalsDisabled;
    }

    /// <summary>The sheet being generated.</summary>
    public IXLWorksheet Worksheet { get; }

    /// <summary>What the range produced, excluding its options slot.</summary>
    public IXLRange GeneratedRange { get; }

    /// <summary>
    /// The range's options slot — its last row when it repeats downwards, its last column when it
    /// repeats across — or <c>null</c> once it has been removed.
    /// </summary>
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
    /// The template expression each line of the range held, keyed by <see cref="OptionTag.Line"/> —
    /// so by worksheet column for a range that repeats downwards, and by row for one that repeats
    /// across. A tag sitting in a line uses it to work out what that line means, which is how
    /// <c>&lt;&lt;Sort&gt;&gt;</c> knows what to sort by without being told twice.
    /// </summary>
    public IReadOnlyDictionary<int, string> LineExpressions { get; }

    /// <summary>
    /// Every tag the range declared, in the order they run.
    /// </summary>
    /// <remarks>
    /// For a tag that has to know what else was asked for. <c>&lt;&lt;Pivot&gt;&gt;</c> reads the
    /// <c>&lt;&lt;Row&gt;&gt;</c> and <c>&lt;&lt;Data&gt;&gt;</c> tags that say what its fields are,
    /// and refuses to run at all when a <c>&lt;&lt;Group&gt;&gt;</c> is present. Includes the tag
    /// doing the reading.
    /// </remarks>
    public IReadOnlyList<OptionTag> Tags { get; }

    /// <summary>Whether the range repeats across rather than downwards.</summary>
    public bool IsHorizontal => Axis.IsHorizontal;

    /// <summary>Which way the range repeats, and the sheet operations that depend on it.</summary>
    internal Ranges.RangeAxis Axis { get; }

    /// <summary>
    /// Pivot tables asked for but not yet built, because they cannot be until every range has
    /// finished moving the sheet about.
    /// </summary>
    internal List<Rewriting.PivotRequest> PendingPivots { get; }

    /// <summary>
    /// Whether the options row should be left without a total, because the template said
    /// <c>&lt;&lt;DisableGrandTotal&gt;&gt;</c>.
    /// </summary>
    /// <remarks>
    /// A grouped range writes the same summary declarations into each group's subtotal row and
    /// again into the options row. A report that wants group totals but no report-wide total sets
    /// this; a summary tag reads it before writing.
    /// </remarks>
    public bool GrandTotalsDisabled { get; }

    /// <summary>
    /// Evaluates a tag parameter as a question and reports whether the answer reads as true.
    /// </summary>
    /// <param name="expression">
    /// A bare expression (<c>item.Quantity &gt; 10</c>), an interpolated one
    /// (<c>{{ ShowWorkings }}</c>), or a literal (<c>true</c>, <c>1</c>) — all three work, because a
    /// template author writes whichever reads best and should not have to know the difference.
    /// </param>
    /// <param name="scope">
    /// The names the expression may use. Defaults to <see cref="Scope"/>; a tag asking about one row
    /// passes that row's scope instead.
    /// </param>
    /// <remarks>
    /// Only <c>null</c> and <c>false</c> are false. Zero and the empty string are <em>true</em>, so a
    /// template meaning "greater than nothing" has to write the comparison out. An expression that
    /// will not evaluate is recorded as an error and answers no, on the principle that a report is
    /// better off leaving a doubtful row out than failing.
    /// </remarks>
    public bool IsTrue(string? expression, ExpressionScope? scope = null)
    {
        if (string.IsNullOrWhiteSpace(expression))
        {
            return false;
        }

        var body = ExpressionText.TryGetSingleExpression(expression, out var single) ? single : expression!;

        try
        {
            return ExpressionTruth.IsTruthy(Engine.Evaluate(body, scope ?? Scope));
        }
        catch (ExpressionEvaluationException ex)
        {
            Errors.Add(new TemplateError(ex.Message, Worksheet.Name, exception: ex));
            return false;
        }
    }
}
