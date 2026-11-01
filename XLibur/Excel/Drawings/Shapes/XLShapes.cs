using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace XLibur.Excel.Drawings;

internal sealed class XLShapes : IXLShapes, IEnumerable<XLShape>
{
    private readonly List<XLShape> _shapes = [];
    private readonly XLWorksheet _worksheet;

    // The shapes holding each name, kept as they change so that a new name is found without reading
    // every shape, as XLPictures does (#804).
    private readonly Dictionary<string, int> _nameUses = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<int, int> _idUses = [];

    internal XLShapes(XLWorksheet worksheet) => _worksheet = worksheet;

    public int Count => _shapes.Count;

    /// <summary>The drawing ids of the deleted shapes the file holds, whose whole anchor the save removes.</summary>
    internal List<int> Deleted { get; } = [];

    /// <summary>Whether a shape was added, deleted or changed since the load or the last save.</summary>
    internal bool HasUnsavedChanges => Deleted.Count > 0 || _shapes.Exists(shape => shape.HasUnsavedChanges);

    public IXLTextBox AddTextBox(string text, IXLCell topLeft)
    {
        ArgumentNullException.ThrowIfNull(text);
        var box = (XLTextBox)Add(XLShapeType.TextBox, topLeft);
        box.Text = text;
        return box;
    }

    public IXLTextBox AddTextBox(string text, int left, int top)
    {
        ArgumentNullException.ThrowIfNull(text);
        var box = (XLTextBox)Add(XLShapeType.TextBox, left, top);
        box.Text = text;
        return box;
    }

    public IXLShape Add(XLShapeType type, IXLCell topLeft)
    {
        ArgumentNullException.ThrowIfNull(topLeft);
        var shape = Create(XLShapeTypes.PresetOf(type), type == XLShapeType.TextBox);
        shape.MoveTo(topLeft);
        return Attach(shape);
    }

    public IXLShape Add(XLShapeType type, int left, int top)
    {
        var shape = Create(XLShapeTypes.PresetOf(type), type == XLShapeType.TextBox);
        shape.MoveTo(left, top);
        return Attach(shape);
    }

    public IXLShape AddPreset(string presetGeometry, IXLCell topLeft)
    {
        ArgumentNullException.ThrowIfNull(presetGeometry);
        ArgumentNullException.ThrowIfNull(topLeft);
        if (!XLShapeTypes.IsPreset(presetGeometry))
            throw new ArgumentException($"'{presetGeometry}' is not a DrawingML preset geometry.", nameof(presetGeometry));

        var shape = Create(presetGeometry, isTextBox: false);
        shape.MoveTo(topLeft);
        return Attach(shape);
    }

    public IXLShape Shape(string name)
    {
        if (TryGetShape(name, out var shape))
            return shape!;

        throw new ArgumentOutOfRangeException(nameof(name), $"Shape {name} was not found.");
    }

    public bool TryGetShape(string name, out IXLShape? shape)
    {
        ArgumentNullException.ThrowIfNull(name);
        shape = _shapes.Find(s => string.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase));
        return shape is not null;
    }

    public bool Contains(string name)
    {
        ArgumentNullException.ThrowIfNull(name);
        return _nameUses.ContainsKey(name);
    }

    public void Delete(string name)
    {
        ArgumentNullException.ThrowIfNull(name);
        var matches = _shapes.Where(s => string.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase)).ToList();
        if (matches.Count == 0)
            throw new ArgumentOutOfRangeException(nameof(name), $"Shape {name} was not found.");

        foreach (var shape in matches)
            Detach(shape);
    }

    public void Delete(IXLShape shape)
    {
        ArgumentNullException.ThrowIfNull(shape);
        if (shape is not XLShape xlShape || !ReferenceEquals(xlShape.Collection, this))
            throw new ArgumentException("The shape is not on this worksheet.", nameof(shape));

        Detach(xlShape);
    }

    public IEnumerator<XLShape> GetEnumerator() => _shapes.GetEnumerator();

    IEnumerator<IXLShape> IEnumerable<IXLShape>.GetEnumerator() => _shapes.Cast<IXLShape>().GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    /// <summary>Whether a shape of this collection is named <paramref name="name"/>, in any case.</summary>
    internal bool IsNameUsed(string name) => _nameUses.ContainsKey(name);

    /// <summary>Whether a shape of this collection has the drawing id <paramref name="id"/>.</summary>
    internal bool IsIdUsed(int id) => _idUses.ContainsKey(id);

    /// <summary>
    /// Adds a shape the loader read from the file, under the name and drawing id the file gives it.
    /// The caller has given a duplicate id a fresh one (#859).
    /// </summary>
    internal XLShape AddLoaded(string presetGeometry, bool isTextBox, string name, int id)
    {
        var shape = isTextBox ? new XLTextBox(_worksheet, presetGeometry) : new XLShape(_worksheet, presetGeometry);
        shape.SetNameUnchecked(name);
        shape.Id = id;
        Attach(shape);
        return shape;
    }

    internal void OnRenamed(XLShape shape, string oldName)
    {
        Release(_nameUses, oldName);
        Take(_nameUses, shape.Name);
    }

    internal void OnIdChanged(int oldId, int newId)
    {
        if (oldId != 0)
            Release(_idUses, oldId);

        if (newId != 0)
            Take(_idUses, newId);
    }

    private XLShape Create(string preset, bool isTextBox)
    {
        var shape = isTextBox ? new XLTextBox(_worksheet, preset) : new XLShape(_worksheet, preset);
        shape.SetNameUnchecked(NextName(XLShapeTypes.NameOf(preset, isTextBox)));
        return shape;
    }

    private XLShape Attach(XLShape shape)
    {
        _shapes.Add(shape);
        shape.Collection = this;
        Take(_nameUses, shape.Name);
        if (shape.Id != 0)
            Take(_idUses, shape.Id);

        return shape;
    }

    private void Detach(XLShape shape)
    {
        _shapes.Remove(shape);
        shape.Collection = null;
        Release(_nameUses, shape.Name);
        if (shape.Id != 0)
            Release(_idUses, shape.Id);

        if (shape.IsInFile)
            Deleted.Add(shape.Id);
    }

    /// <summary>The first <c>"{start} n"</c>, counting from one more than the shapes there are, that no shape or picture has.</summary>
    private string NextName(string start)
    {
        var pictures = (XLPictures)_worksheet.Pictures;
        for (var number = _shapes.Count + 1; ; number++)
        {
            var name = $"{start} {number}";
            if (!IsNameUsed(name) && !pictures.IsNameTaken(name))
                return name;
        }
    }

    private static void Take<TKey>(Dictionary<TKey, int> uses, TKey key) where TKey : notnull =>
        uses[key] = uses.GetValueOrDefault(key) + 1;

    private static void Release<TKey>(Dictionary<TKey, int> uses, TKey key) where TKey : notnull
    {
        if (!uses.TryGetValue(key, out var count))
            return;

        if (count > 1)
            uses[key] = count - 1;
        else
            uses.Remove(key);
    }
}
