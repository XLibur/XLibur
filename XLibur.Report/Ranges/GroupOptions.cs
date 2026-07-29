using System;
using System.Collections.Generic;
using XLibur.Report.Tags;

namespace XLibur.Report.Ranges;

/// <summary>
/// The grouping options a range set for all of its levels at once.
/// </summary>
/// <remarks>
/// A <see cref="GroupOptionTag"/> carries no behaviour of its own — it is read here, before the
/// first row is written, because what it says has to be settled before the rows it affects exist.
/// A level may still turn an option on for itself alone, so these are floors, never ceilings.
/// </remarks>
internal sealed class GroupOptions
{
    public bool SummaryAbove { get; private set; }

    public bool MergeLabels { get; private set; }

    public bool PageBreaks { get; private set; }

    public bool Collapse { get; private set; }

    public bool NoSubtotals { get; private set; }

    public bool GrandTotalDisabled { get; private set; }

    public static GroupOptions Read(IEnumerable<OptionTag> tags)
    {
        var options = new GroupOptions();

        foreach (var tag in tags)
        {
            if (tag is not GroupOptionTag)
            {
                continue;
            }

            var name = tag.Token.Name;

            if (Is(name, "SummaryAbove"))
            {
                options.SummaryAbove = true;
            }
            else if (Is(name, "MergeLabels"))
            {
                options.MergeLabels = true;
            }
            else if (Is(name, "PageBreaks"))
            {
                options.PageBreaks = true;
            }
            else if (Is(name, "Collapse"))
            {
                options.Collapse = true;
            }
            else if (Is(name, "DisableSubtotals"))
            {
                options.NoSubtotals = true;
            }
            else if (Is(name, "DisableGrandTotal"))
            {
                options.GrandTotalDisabled = true;
            }
        }

        return options;
    }

    private static bool Is(string name, string option) =>
        string.Equals(name, option, StringComparison.OrdinalIgnoreCase);
}
