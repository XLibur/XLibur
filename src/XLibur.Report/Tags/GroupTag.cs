namespace XLibur.Report.Tags;

/// <summary>
/// Groups a range's rows by the column it is written in, giving each group a subtotal row and an
/// Excel outline level. Written in the options row as <c>&lt;&lt;Group&gt;&gt;</c>.
/// </summary>
/// <remarks>
/// <para>
/// The group key is the template expression of the column the tag sits under, the same way
/// <see cref="SortTag"/> takes its sort key, so a column already showing <c>{{ item.Region }}</c>
/// needs no second mention of it. Give <c>by</c> to group by something the range does not display.
/// </para>
/// <para>
/// Rows are ordered by their group keys before anything is written, so groups come out contiguous
/// whatever order the data arrived in. That ordering is stable, so a <c>&lt;&lt;Sort&gt;&gt;</c> on
/// another column still decides the order <em>within</em> a group. Give <c>nosort</c> to leave the
/// order alone and group runs of equal values as they come.
/// </para>
/// <para>
/// Several <c>&lt;&lt;Group&gt;&gt;</c> tags nest, leftmost outermost. Each group's subtotal row
/// repeats whatever summary tags — <c>&lt;&lt;Sum&gt;&gt;</c> and friends — the options row
/// declares, over that group's rows alone, and takes the options row's styling so that a subtotal
/// and the grand total look alike.
/// </para>
/// <para>
/// Parameters: <c>by</c> (group key expression), <c>desc</c> (largest group first), <c>nosort</c>
/// (group the given order as-is), <c>totalLabel</c> (the subtotal row's label, where <c>{0}</c>
/// stands for the group key — <c>"{0} Total"</c> by default), <c>merge</c> or <c>mergeLabels</c>
/// (merge the group's label cells into one), <c>summaryAbove</c> (subtotal above the group rather
/// than below), <c>pageBreaks</c> (a page break after each group), <c>collapse</c> (write the
/// outline collapsed), <c>disableSubtotals</c> (outline the group without a subtotal row).
/// </para>
/// </remarks>
public sealed class GroupTag : OptionTag
{
    /// <summary>The expression to group by, or an empty string to use the column's own.</summary>
    internal string KeyExpression => Token.Value("by");

    /// <summary>Whether groups run largest key first.</summary>
    internal bool Descending => Token.Flag("desc");

    /// <summary>Whether the given row order is kept, grouping only runs of equal keys.</summary>
    internal bool KeepOrder => Token.Flag("nosort");

    /// <summary>Whether the group's label cells are merged into one.</summary>
    internal bool MergeLabels => Token.Flag("mergeLabels") || Token.Flag("merge");

    /// <summary>Whether the subtotal row sits above the group rather than below it.</summary>
    internal bool SummaryAbove => Token.Flag("summaryAbove");

    /// <summary>Whether a page break follows each group.</summary>
    internal bool PageBreaks => Token.Flag("pageBreaks");

    /// <summary>Whether the outline is written collapsed.</summary>
    internal bool Collapse => Token.Flag("collapse");

    /// <summary>Whether the level is outlined without a subtotal row.</summary>
    internal bool NoSubtotals => Token.Flag("disableSubtotals") || Token.Flag("noSubtotals");

    /// <summary>The subtotal row's label, with <c>{0}</c> standing for the group key.</summary>
    internal string TotalLabel => Token.Value("totalLabel", "{0} Total");
}

/// <summary>
/// Sets a grouping option for a whole range rather than for one <c>&lt;&lt;Group&gt;&gt;</c> level.
/// Written anywhere in the options row.
/// </summary>
/// <remarks>
/// <para>
/// The options are <c>&lt;&lt;SummaryAbove&gt;&gt;</c>, <c>&lt;&lt;MergeLabels&gt;&gt;</c>,
/// <c>&lt;&lt;PageBreaks&gt;&gt;</c>, <c>&lt;&lt;Collapse&gt;&gt;</c>,
/// <c>&lt;&lt;DisableSubtotals&gt;&gt;</c> and <c>&lt;&lt;DisableGrandTotal&gt;&gt;</c>. Each is
/// the range-wide form of the like-named <c>&lt;&lt;Group&gt;&gt;</c> parameter, saving a template
/// with several group levels from repeating itself; a level may still turn one on for itself
/// alone.
/// </para>
/// <para>
/// <c>&lt;&lt;DisableGrandTotal&gt;&gt;</c> is the exception, having no per-level form: it leaves
/// the options row's own summaries unwritten, which is how a report shows a total per group and
/// none for the report.
/// </para>
/// <para>
/// The tag does nothing when it runs. It is read before the first row is written, because what it
/// says has to be settled before the rows it affects exist.
/// </para>
/// </remarks>
public sealed class GroupOptionTag : OptionTag
{
}
