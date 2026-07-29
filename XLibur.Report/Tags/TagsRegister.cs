using System;
using System.Collections.Concurrent;
using System.Collections.Generic;

namespace XLibur.Report.Tags;

/// <summary>
/// The tags templates may use. Built-in tags are registered on first use; add your own with
/// <see cref="Add{T}"/>.
/// </summary>
public static class TagsRegister
{
    private static readonly ConcurrentDictionary<string, Registration> Registrations =
        new(StringComparer.OrdinalIgnoreCase);

    static TagsRegister() => RegisterBuiltIns();

    /// <summary>
    /// Registers <typeparamref name="T"/> under <paramref name="name"/>, replacing any tag already
    /// registered under it.
    /// </summary>
    /// <param name="name">The name a template writes, without the angle brackets.</param>
    /// <param name="priority">
    /// Lower runs first. Use it when one tag has to see what another did — a total has to be
    /// written before a column is fitted to its contents.
    /// </param>
    public static void Add<T>(string name, byte priority = 128)
        where T : OptionTag, new()
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            throw new ArgumentException("A tag name is required.", nameof(name));
        }

        Registrations[name] = new Registration(priority, () => new T());
    }

    /// <summary>Whether a tag is registered under <paramref name="name"/>.</summary>
    public static bool Contains(string name) => Registrations.ContainsKey(name);

    /// <summary>The names of every registered tag.</summary>
    public static IEnumerable<string> Names => Registrations.Keys;

    internal static bool TryCreate(TagToken token, int row, int column, int line, out OptionTag tag)
    {
        if (!Registrations.TryGetValue(token.Name, out var registration))
        {
            tag = null!;
            return false;
        }

        tag = registration.Create();
        tag.Token = token;
        tag.Row = row;
        tag.Column = column;
        tag.Line = line;
        return true;
    }

    internal static byte PriorityOf(string name) =>
        Registrations.TryGetValue(name, out var registration) ? registration.Priority : (byte)128;

    private static void RegisterBuiltIns()
    {
        // Ordering matters where one tag reads what another wrote: totals land before the
        // column-fitting and autofilter tags measure the block.
        // Read before the range is even parsed, not run: which way the range repeats decides where
        // its options slot is, so the tag cannot wait its turn.
        Add<HorizontalTag>("Horizontal", 0);

        Add<GroupOptionTag>("SummaryAbove", 5);
        Add<GroupOptionTag>("MergeLabels", 5);
        Add<GroupOptionTag>("PageBreaks", 5);
        Add<GroupOptionTag>("Collapse", 5);
        Add<GroupOptionTag>("DisableSubtotals", 5);
        Add<GroupOptionTag>("DisableGrandTotal", 5);

        // Before everything: what a test leaves out should not be sorted, grouped or totalled, and
        // nothing downstream has to know anything was dropped.
        Add<IfTag>("If", 1);

        Add<SortTag>("Sort", 10);
        Add<SortTag>("Asc", 10);
        Add<DescendingSortTag>("Desc", 10);

        // Grouping reads the row order sorting left behind, and nests its levels left to right —
        // which the register preserves, ordering equal priorities by the column they were read in.
        Add<GroupTag>("Group", 20);

        Add<SummaryFunctionTag>("Sum", 50);
        Add<SummaryFunctionTag>("Avg", 50);
        Add<SummaryFunctionTag>("Average", 50);
        Add<SummaryFunctionTag>("Count", 50);
        Add<SummaryFunctionTag>("CountA", 50);
        Add<SummaryFunctionTag>("Max", 50);
        Add<SummaryFunctionTag>("Min", 50);
        Add<SummaryFunctionTag>("Product", 50);
        Add<SummaryFunctionTag>("StdDev", 50);
        Add<SummaryFunctionTag>("StdDevP", 50);
        Add<SummaryFunctionTag>("Var", 50);
        Add<SummaryFunctionTag>("VarP", 50);

        // Declarations rather than actions: <<Pivot>> reads them, and they report a template mistake
        // if there is no <<Pivot>> to read them.
        Add<PivotFieldTag>("Row", 200);
        Add<PivotFieldTag>("Column", 200);
        Add<PivotFieldTag>("Col", 200);
        Add<PivotFieldTag>("Page", 200);
        Add<PivotFieldTag>("Data", 200);

        Add<AutoFilterTag>("AutoFilter", 200);
        Add<ColumnsFitTag>("ColsFit", 200);
        Add<RowsFitTag>("RowsFit", 200);
        Add<HiddenTag>("Hidden", 200);
        Add<HiddenTag>("Hide", 200);
        Add<DeleteTag>("Delete", 250);

        // Last of all: a pivot cache's source is a plain rectangle, so the pivot has to be built over
        // the columns that survive <<Delete>> rather than over the ones the template drew.
        Add<PivotTag>("Pivot", 255);
    }

    private sealed record Registration(byte Priority, Func<OptionTag> Create);
}
