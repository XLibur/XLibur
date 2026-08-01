using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using XLibur.Excel;

namespace XLibur.Report.Ranges;

/// <summary>
/// Works out which of a workbook's defined names mark repeating template rows.
/// </summary>
/// <remarks>
/// <para>
/// A name marks a repeating range when it resolves to a collection: either a variable of the same
/// name, or — for a name like <c>Customer_Orders</c> — a property path walked from one. Names that
/// resolve to nothing, or to a single value, are left alone; a template is free to use defined
/// names for their ordinary purpose.
/// </para>
/// <para>
/// Names are gathered under Excel's scoping rules rather than de-duplicated workbook-wide, so a
/// template may declare the same name once per sheet — <c>Q1!Items</c> and <c>Q2!Items</c> — and
/// have every one of them bind. See <see cref="EnumerateNames"/>.
/// </para>
/// <para>
/// A name is then matched to its variable the way Excel matches names, without regard to case. That
/// is the boundary at which an Excel identifier becomes a C# one, and matching stops being
/// case-insensitive on the other side of it. See <see cref="TryGetVariable"/>.
/// </para>
/// </remarks>
internal static class RangeBinder
{
    /// <summary>
    /// Resolves every bindable defined name in <paramref name="workbook"/>, ordered by sheet and
    /// then top to bottom so expansion runs down the sheet.
    /// </summary>
    public static List<BoundRange> Resolve(
        IXLWorkbook workbook,
        IReadOnlyDictionary<string, object?> variables,
        TemplateErrors errors)
    {
        var bound = new List<BoundRange>();

        foreach (var definedName in EnumerateNames(workbook))
        {
            if (!TryResolveItems(definedName.Name, variables, errors, out var items))
            {
                continue;
            }

            var ranges = definedName.Ranges.ToList();
            if (ranges.Count == 0)
            {
                continue;
            }

            if (ranges.Count > 1)
            {
                errors.Add(new TemplateError(
                    $"Defined name '{definedName.Name}' covers {ranges.Count} areas; only single-area names can be bound to data.",
                    ranges[0].Worksheet.Name));
                continue;
            }

            var range = ranges[0];
            bound.Add(new BoundRange(definedName, range.Worksheet, items!, RangeArea.From(range)));
        }

        return bound
            .OrderBy(b => b.Worksheet.Position)
            .ThenBy(b => b.Area.FirstRow)
            .ThenBy(b => b.Area.FirstColumn)
            .ToList();
    }

    /// <summary>
    /// The names a template can bind, with Excel's scoping applied.
    /// </summary>
    /// <remarks>
    /// <para>
    /// Every sheet-scoped name is yielded — the same name on two sheets is two names in Excel, and a
    /// template with one section per sheet is the obvious way to write it — followed by each
    /// workbook-scoped name that no sheet has shadowed.
    /// </para>
    /// <para>
    /// Excel decides shadowing by where a reference is written: a sheet-scoped name hides the
    /// workbook-scoped one of that name on its own sheet, and nowhere else. A defined name here is
    /// not read from anywhere, it <em>is</em> a range, so the sheet it sits on stands in for the
    /// sheet a reference would be written on. A workbook-scoped <c>Items</c> covering <c>Q1!A5:C5</c>
    /// therefore keeps binding when <c>Q2</c> declares its own <c>Items</c>, and is dropped when
    /// <c>Q1</c> does.
    /// </para>
    /// </remarks>
    private static IEnumerable<IXLDefinedName> EnumerateNames(IXLWorkbook workbook)
    {
        var declaringSheets = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);

        foreach (var worksheet in workbook.Worksheets)
        {
            foreach (var name in worksheet.DefinedNames.ValidNamedRanges())
            {
                if (!declaringSheets.TryGetValue(name.Name, out var sheets))
                {
                    sheets = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    declaringSheets[name.Name] = sheets;
                }

                sheets.Add(worksheet.Name);
                yield return name;
            }
        }

        foreach (var name in workbook.DefinedNames.ValidNamedRanges())
        {
            if (!IsShadowed(name, declaringSheets))
            {
                yield return name;
            }
        }
    }

    /// <summary>
    /// Whether the sheet a workbook-scoped name covers declares a name of its own by that name.
    /// </summary>
    /// <remarks>
    /// The name's ranges are only read when some sheet declares the name at all, which is almost
    /// never; a lookup that misses costs a dictionary probe and no parsing of <c>RefersTo</c>.
    /// </remarks>
    private static bool IsShadowed(IXLDefinedName name, Dictionary<string, HashSet<string>> declaringSheets)
    {
        if (!declaringSheets.TryGetValue(name.Name, out var sheets))
        {
            return false;
        }

        return name.Ranges.Any(range => sheets.Contains(range.Worksheet.Name));
    }

    /// <summary>
    /// Resolves a defined name to a collection, if it names one.
    /// </summary>
    private static bool TryResolveItems(
        string name,
        IReadOnlyDictionary<string, object?> variables,
        TemplateErrors errors,
        out IReadOnlyList<object?>? items)
    {
        items = null;

        if (!TryResolveValue(name, variables, errors, out var value))
        {
            return false;
        }

        // A string is enumerable but is never what a template author means by a data range.
        if (value is null or string || value is not IEnumerable enumerable)
        {
            return false;
        }

        items = enumerable.Cast<object?>().ToList();
        return true;
    }

    private static bool TryResolveValue(
        string name,
        IReadOnlyDictionary<string, object?> variables,
        TemplateErrors errors,
        out object? value)
    {
        if (TryGetVariable(name, name, variables, errors, out value))
        {
            return true;
        }

        // Underscores separate a property path: Customer_Orders binds Customer.Orders.
        if (name.IndexOf('_') < 0)
        {
            return false;
        }

        var segments = name.Split('_');
        if (!TryGetVariable(name, segments[0], variables, errors, out var current) || current is null)
        {
            return false;
        }

        for (var i = 1; i < segments.Length; i++)
        {
            current = ReadMember(current, segments[i]);
            if (current is null)
            {
                return false;
            }
        }

        value = current;
        return true;
    }

    /// <summary>Finds the variable a defined name selects.</summary>
    /// <param name="definedName">The whole name, for the error message.</param>
    /// <param name="name">The name to look up — the whole of it, or a path's first segment.</param>
    /// <param name="variables">The template's variables.</param>
    /// <param name="errors">Collects an ambiguous match.</param>
    /// <param name="value">The variable's value, when one was found.</param>
    /// <remarks>
    /// <para>
    /// Excel holds defined names in one case-insensitive namespace, so a template author who types
    /// <c>ITEMS</c> in the name box has named the same thing as the variable added as <c>Items</c>,
    /// and it binds. An exact match still wins, so nothing that binds today binds anything
    /// different; the case-insensitive pass runs only when there is nothing exact to find.
    /// </para>
    /// <para>
    /// The boundary stops here. The property segments of <c>Customer_Orders</c> are read by
    /// <see cref="ReadMember"/>, and every name inside <c>{{ }}</c> by the expression engine, as the
    /// C# members they are: <c>item.Price</c> and <c>item.price</c> stay different names.
    /// </para>
    /// <para>
    /// Two variables differing only by case are the one case with no answer — nothing is exact, and
    /// picking either would be picking whichever the dictionary happened to yield first — so the
    /// name is reported and left unbound.
    /// </para>
    /// </remarks>
    private static bool TryGetVariable(
        string definedName,
        string name,
        IReadOnlyDictionary<string, object?> variables,
        TemplateErrors errors,
        out object? value)
    {
        if (variables.TryGetValue(name, out value))
        {
            return true;
        }

        string? matched = null;

        foreach (var variable in variables)
        {
            if (!StringComparer.OrdinalIgnoreCase.Equals(variable.Key, name))
            {
                continue;
            }

            if (matched is not null)
            {
                // Say which part of the name did the matching, because for a property path it is the
                // first segment rather than the whole of it.
                var through = string.Equals(definedName, name, StringComparison.Ordinal)
                    ? string.Empty
                    : $" (through its first segment '{name}')";

                errors.Add(new TemplateError(
                    $"Defined name '{definedName}'{through} matches the variables '{matched}' and " +
                    $"'{variable.Key}', which differ only by case. Rename one of them, or spell the name to " +
                    "match one of them exactly."));
                value = null;
                return false;
            }

            matched = variable.Key;
            value = variable.Value;
        }

        return matched is not null;
    }

    /// <summary>Reads a property, field or dictionary entry by name.</summary>
    internal static object? ReadMember(object? target, string name)
    {
        if (target is null)
        {
            return null;
        }

        if (target is IDictionary dictionary)
        {
            return dictionary.Contains(name) ? dictionary[name] : null;
        }

        var type = target.GetType();

        var property = type.GetProperty(name, BindingFlags.Public | BindingFlags.Instance);
        if (property is not null && property.CanRead && property.GetIndexParameters().Length == 0)
        {
            return property.GetValue(target);
        }

        var field = type.GetField(name, BindingFlags.Public | BindingFlags.Instance);
        return field?.GetValue(target);
    }
}
