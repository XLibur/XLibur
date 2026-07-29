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
/// A name marks a repeating range when it resolves to a collection: either a variable of the same
/// name, or — for a name like <c>Customer_Orders</c> — a property path walked from one. Names that
/// resolve to nothing, or to a single value, are left alone; a template is free to use defined
/// names for their ordinary purpose.
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
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var definedName in EnumerateNames(workbook))
        {
            if (!seen.Add(definedName.Name))
            {
                continue;
            }

            if (!TryResolveItems(definedName.Name, variables, out var items))
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

    private static IEnumerable<IXLDefinedName> EnumerateNames(IXLWorkbook workbook)
    {
        foreach (var name in workbook.DefinedNames.ValidNamedRanges())
        {
            yield return name;
        }

        foreach (var worksheet in workbook.Worksheets)
        {
            foreach (var name in worksheet.DefinedNames.ValidNamedRanges())
            {
                yield return name;
            }
        }
    }

    /// <summary>
    /// Resolves a defined name to a collection, if it names one.
    /// </summary>
    private static bool TryResolveItems(
        string name,
        IReadOnlyDictionary<string, object?> variables,
        out IReadOnlyList<object?>? items)
    {
        items = null;

        if (!TryResolveValue(name, variables, out var value))
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
        out object? value)
    {
        if (variables.TryGetValue(name, out value))
        {
            return true;
        }

        // Underscores separate a property path: Customer_Orders binds Customer.Orders.
        if (name.IndexOf('_') < 0)
        {
            return false;
        }

        var segments = name.Split('_');
        if (!variables.TryGetValue(segments[0], out var current) || current is null)
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
