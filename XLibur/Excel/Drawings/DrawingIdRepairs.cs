using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace XLibur.Excel.Drawings;

/// <summary>
/// The drawing ids a load gave the pictures and shapes of a sheet whose id another one of their
/// kind already had (#859), for the next save to write into the drawing.
/// </summary>
/// <remarks>
/// <para>
/// The schema makes a drawing id unique within a drawing, but files break the rule — copy-paste bugs,
/// other producers — and Excel opens them. The save finds a picture by its id and image link and a
/// shape by its id, so a second one under the same id could take the first one's edits. The load
/// gives each later one an id of its own, as Excel does when it repairs the file, and the save
/// writes it in. A connector that names the duplicate by id then names the first one.
/// </para>
/// <para>
/// A repair names its element by position: the n-th element of its kind with the old id, in
/// document order. The save starts from the file the load read, or from what the last save wrote,
/// so the drawing is as the load saw it until the repairs are written, which the save does before
/// anything else. All of one save's repairs are found before any is written, so writing one does
/// not move another. They are dropped once that save has succeeded.
/// </para>
/// </remarks>
internal sealed class DrawingIdRepairs
{
    internal enum Kind
    {
        /// <summary>An <c>xdr:pic</c>, anywhere in the drawing.</summary>
        Picture,

        /// <summary>An <c>xdr:sp</c> that an anchor holds directly, as the loader reads shapes.</summary>
        Shape,
    }

    // Created by the first repair: a sheet whose drawing ids are all its own, nearly every sheet, has none.
    private List<(Kind Kind, uint OldId, int Occurrence, uint NewId)>? _repairs;
    private uint _highest;

    // The drawing the load is reading, indexed once by kind and id rather than walked once per
    // duplicate, as #797 and #804 did for the save.
    private Xdr.WorksheetDrawing? _indexed;
    private Dictionary<(Kind Kind, uint Id), List<OpenXmlElement>>? _index;

    internal int Count => _repairs?.Count ?? 0;

    /// <summary>
    /// Takes a fresh id for <paramref name="element"/>, whose id <paramref name="oldId"/> another of
    /// its kind in <paramref name="drawing"/> already has, or which has no id: one above every id of
    /// the drawing and every id given before.
    /// </summary>
    internal int Repair(Xdr.WorksheetDrawing drawing, Kind kind, OpenXmlElement element, uint oldId)
    {
        if (_indexed != drawing)
        {
            _indexed = drawing;
            _index = Index(drawing);
            _highest = Math.Max(_highest, HighestId(drawing));
        }

        var occurrence = _index!.TryGetValue((kind, oldId), out var elements) ? elements.IndexOf(element) : -1;
        var newId = ++_highest;
        (_repairs ??= []).Add((kind, oldId, occurrence, newId));
        return checked((int)newId);
    }

    /// <summary>Writes every repair into <paramref name="drawing"/>, the drawing as the load read it.</summary>
    internal void Apply(Xdr.WorksheetDrawing drawing)
    {
        if (_repairs is null)
            return;

        var index = Index(drawing);
        var found = new List<(Xdr.NonVisualDrawingProperties Properties, uint NewId)>();
        foreach (var (kind, oldId, occurrence, newId) in _repairs)
        {
            if (index.TryGetValue((kind, oldId), out var elements) &&
                occurrence >= 0 && occurrence < elements.Count &&
                PropertiesOf(elements[occurrence]) is { } properties)
                found.Add((properties, newId));
        }

        foreach (var (properties, newId) in found)
            properties.Id = newId;
    }

    /// <summary>Forgets the repairs, once a save that wrote them has succeeded.</summary>
    internal void Clear()
    {
        _repairs = null;
        _highest = 0;
        _indexed = null;
        _index = null;
    }

    /// <summary>The pictures and shapes of the drawing by kind and id, each list in document order.</summary>
    private static Dictionary<(Kind Kind, uint Id), List<OpenXmlElement>> Index(Xdr.WorksheetDrawing drawing)
    {
        var index = new Dictionary<(Kind Kind, uint Id), List<OpenXmlElement>>();
        foreach (var picture in drawing.Descendants<Xdr.Picture>())
            Add(index, Kind.Picture, picture);

        foreach (var anchor in drawing.ChildElements)
        {
            foreach (var shape in anchor.Elements<Xdr.Shape>())
                Add(index, Kind.Shape, shape);
        }

        return index;
    }

    private static void Add(Dictionary<(Kind Kind, uint Id), List<OpenXmlElement>> index, Kind kind, OpenXmlElement element)
    {
        var id = PropertiesOf(element)?.Id?.Value ?? 0;
        if (!index.TryGetValue((kind, id), out var elements))
            index[(kind, id)] = elements = [];

        elements.Add(element);
    }

    private static Xdr.NonVisualDrawingProperties? PropertiesOf(OpenXmlElement element) => element switch
    {
        Xdr.Picture picture => picture.NonVisualPictureProperties?.NonVisualDrawingProperties,
        Xdr.Shape shape => shape.NonVisualShapeProperties?.NonVisualDrawingProperties,
        _ => null,
    };

    private static uint HighestId(Xdr.WorksheetDrawing drawing)
    {
        uint highest = 0;
        foreach (var properties in drawing.Descendants<Xdr.NonVisualDrawingProperties>())
            highest = Math.Max(highest, properties.Id?.Value ?? 0);

        return highest;
    }
}
