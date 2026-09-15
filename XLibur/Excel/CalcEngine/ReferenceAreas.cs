using System;
using System.Collections.Generic;
using System.Diagnostics;
using XLibur.Excel.Coordinates;

namespace XLibur.Excel.CalcEngine;

/// <summary>
/// What a node gives <see cref="DependenciesVisitor"/> back: either "not a reference" (the default
/// value), or a reference to zero or more sheet areas.
/// </summary>
/// <remarks>
/// One area is the usual case. The struct holds it in a field, so most nodes allocate nothing. A list
/// is created only when a union, an <c>IF</c> or a <c>CHOOSE</c> joins references. Before, the visitor
/// returned a new <c>List&lt;SheetArea&gt;</c> for each reference node (#513).
/// </remarks>
internal readonly struct ReferenceAreas
{
    private readonly SheetArea _single;
    private readonly List<SheetArea>? _many;
    private readonly bool _hasSingle;

    private ReferenceAreas(SheetArea single)
    {
        _single = single;
        _hasSingle = true;
        IsReference = true;
    }

    private ReferenceAreas(List<SheetArea> many)
    {
        _many = many;
        IsReference = true;
    }

    /// <summary>
    /// Not a reference: a value, an array or an error.
    /// </summary>
    public static ReferenceAreas None => default;

    /// <summary>
    /// A reference to one area.
    /// </summary>
    public static ReferenceAreas Of(SheetArea area) => new(area);

    /// <summary>
    /// A reference to the areas in <paramref name="areas"/>, which can be empty. The result owns the
    /// list from then on.
    /// </summary>
    public static ReferenceAreas Of(List<SheetArea> areas) => new(areas);

    /// <summary>
    /// Is this a reference? A reference can have no areas, for example a range between two sheets.
    /// </summary>
    public bool IsReference { get; }

    /// <summary>
    /// The number of areas. <c>0</c> when this is not a reference.
    /// </summary>
    public int Count => _many?.Count ?? (_hasSingle ? 1 : 0);

    public SheetArea this[int index]
    {
        get
        {
            if (_many is not null)
                return _many[index];

            if (index != 0 || !_hasSingle)
                throw new ArgumentOutOfRangeException(nameof(index));

            return _single;
        }
    }

    /// <summary>
    /// A reference to the areas of this reference and of <paramref name="other"/>. Both must be
    /// references.
    /// </summary>
    /// <remarks>
    /// When this value already holds a list, the list grows in place. Only the value that created a
    /// list holds it, and the visitor does not use a value again after it joins it with another, so
    /// no other value sees the change. The visitor did the same with <c>List.AddRange</c> before.
    /// </remarks>
    public ReferenceAreas Concat(ReferenceAreas other)
    {
        Debug.Assert(IsReference && other.IsReference, "Only references are joined.");
        if (other.Count == 0)
            return this;

        if (Count == 0)
            return other;

        var areas = _many ?? new List<SheetArea>(1 + other.Count) { _single };
        for (var i = 0; i < other.Count; ++i)
            areas.Add(other[i]);

        return new ReferenceAreas(areas);
    }
}
