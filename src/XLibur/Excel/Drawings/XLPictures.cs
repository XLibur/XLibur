using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace XLibur.Excel.Drawings;

internal sealed class XLPictures : IXLPictures, IEnumerable<XLPicture>
{
    private readonly List<XLPicture> _pictures = [];
    private readonly XLWorksheet _worksheet;
    private long _nextGroupKey = 1;
    private HashSet<string> _keptShapeNames = new(StringComparer.OrdinalIgnoreCase);

    // The pictures holding each name and each id, kept as the pictures change so that adding a
    // picture does not read every picture (#804). A file can give two pictures one name or one id.
    private readonly Dictionary<string, int> _nameUses = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<int, int> _idUses = [];

    // The highest id, and the highest id below it, or null when not known. The one below is the
    // highest once the picture holding the highest gives it up, as a loaded picture does: it takes
    // the next free id, the highest, when it is created, then the file's. Without it each loaded
    // picture lost the highest, and the next one read every id to find it again (#804).
    private int? _highestId;
    private int? _belowHighestId;

    public XLPictures(XLWorksheet worksheet)
    {
        _worksheet = worksheet;
        Deleted = new List<(int Id, string RelId)>();
        DeletedFromGroups = new List<(int Id, string? RelId)>();
        PendingGroups = new List<XLPendingGroup>();
    }

    public int Count
    {
        [DebuggerStepThrough]
        get => _pictures.Count;
    }

    /// <summary>
    /// Deleted top-level pictures, whose whole anchor is removed on save. Keyed by drawing id plus
    /// rel id, because pictures that share an image part have the same rel id.
    /// </summary>
    internal ICollection<(int Id, string RelId)> Deleted { get; private set; }

    /// <summary>
    /// Deleted pictures that lived inside a group shape. Keyed by drawing id (to locate the exact
    /// <c>xdr:pic</c>) plus rel id (to drop the image part if unreferenced). Removed in place on save
    /// so the group and its other shapes survive.
    /// </summary>
    internal ICollection<(int Id, string? RelId)> DeletedFromGroups { get; }

    /// <summary>Groups to create on the next save (members are moved into a new group shape).</summary>
    internal ICollection<XLPendingGroup> PendingGroups { get; }

    public IXLPicture Add(Stream stream)
    {
        var picture = new XLPicture(_worksheet, stream);
        Attach(picture);
        picture.Name = GetNextPictureName();
        return picture;
    }

    public IXLPicture Add(Stream stream, string name)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(name);
        ThrowIfNameTaken(name);
        var picture = Add(stream);
        picture.Name = name;
        return picture;
    }

    public IXLPicture Add(Stream stream, XLPictureFormat format)
    {
        var picture = new XLPicture(_worksheet, stream, format);
        Attach(picture);
        picture.Name = GetNextPictureName();
        return picture;
    }

    public IXLPicture Add(Stream stream, XLPictureFormat format, string name)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(name);
        ThrowIfNameTaken(name);
        var picture = Add(stream, format);
        picture.Name = name;
        return picture;
    }

    public IXLPicture Add(string imageFile)
    {
        ArgumentException.ThrowIfNullOrEmpty(imageFile);
        using var fs = File.OpenRead(imageFile);
        var picture = new XLPicture(_worksheet, fs);
        Attach(picture);
        picture.Name = GetNextPictureName();
        return picture;
    }

    public IXLPicture Add(string imageFile, string name)
    {
        ArgumentException.ThrowIfNullOrEmpty(imageFile);
        ArgumentException.ThrowIfNullOrWhiteSpace(name);
        ThrowIfNameTaken(name);
        var picture = Add(imageFile);
        picture.Name = name;
        return picture;
    }

    public bool Contains(string pictureName)
    {
        ArgumentNullException.ThrowIfNull(pictureName);
        return _pictures.Any(p => string.Equals(p.Name, pictureName, StringComparison.OrdinalIgnoreCase));
    }

    public void Delete(IXLPicture picture)
    {
        Delete(picture.Name);
    }

    public void Delete(string pictureName)
    {
        ArgumentException.ThrowIfNullOrEmpty(pictureName);
        var picturesToDelete = _pictures
            .Where(picture => picture.Name.Equals(pictureName, StringComparison.OrdinalIgnoreCase))
            .ToList();

        if (picturesToDelete.Count == 0)
            throw new ArgumentOutOfRangeException(nameof(pictureName), $"Picture {pictureName} was not found.");

        foreach (var picture in picturesToDelete)
        {
            if (picture.IsInGroup)
                DeletedFromGroups.Add((picture.Id, picture.RelId));
            else if (!string.IsNullOrEmpty(picture.RelId))
                Deleted.Add((picture.Id, picture.RelId));

            Detach(picture);
        }
    }

    IEnumerator<IXLPicture> IEnumerable<IXLPicture>.GetEnumerator()
    {
        return _pictures.Cast<IXLPicture>().GetEnumerator();
    }

    public IEnumerator<XLPicture> GetEnumerator()
    {
        return ((IEnumerable<XLPicture>)_pictures).GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    public IXLPicture Picture(string pictureName)
    {
        ArgumentException.ThrowIfNullOrEmpty(pictureName);
        if (TryGetPicture(pictureName, out IXLPicture? p))
            return p!;

        throw new ArgumentOutOfRangeException(nameof(pictureName), $"Picture {pictureName} was not found.");
    }

    public bool TryGetPicture(string pictureName, out IXLPicture? picture)
    {
        ArgumentNullException.ThrowIfNull(pictureName);
        var match = _pictures.FirstOrDefault(p => p.Name.Equals(pictureName, StringComparison.OrdinalIgnoreCase));
        if (match is not null)
        {
            picture = match;
            return true;
        }

        picture = null;
        return false;
    }

    /// <summary>
    /// Adds a loaded picture under the name and id the file gives it. No name is generated: while a
    /// drawing loads every name in it is taken (see <see cref="LoadDrawing"/>), so a generated name
    /// would walk past all of them only to be replaced.
    /// </summary>
    internal IXLPicture Add(Stream stream, string name, int id)
    {
        var picture = new XLPicture(_worksheet, stream);
        Attach(picture);
        picture.SetName(name);
        picture.Id = id;
        return picture;
    }

    /// <summary>
    /// Add a picture as a new member of the same group as <paramref name="groupMember"/>. The new
    /// picture inherits the group's composed transform, so its <c>Width</c>/<c>Height</c>/<c>Left</c>/
    /// <c>Top</c> are interpreted in sheet-space; the writer creates the <c>xdr:pic</c> inside the
    /// group on save. (A first-class group API is added in a later phase.)
    /// </summary>
    internal XLPicture AddToGroup(XLPicture groupMember, Stream stream, string? name = null)
    {
        var source = groupMember.GroupInfo
                     ?? throw new ArgumentException("The picture is not part of a group.", nameof(groupMember));

        var picture = (XLPicture)(name is null ? Add(stream) : Add(stream, name));
        picture.Placement = XLPicturePlacement.FreeFloating;
        picture.GroupInfo = new XLPictureGroup
        {
            ScaleX = source.ScaleX,
            ScaleY = source.ScaleY,
            OffsetX = source.OffsetX,
            OffsetY = source.OffsetY,
            GroupId = source.GroupId,
            GroupKey = source.GroupKey,
            IsNew = true,
        };
        return picture;
    }

    /// <summary>Allocate a stable, per-worksheet identity for a group (see <see cref="XLPictureGroup.GroupKey"/>).</summary>
    internal long AllocateGroupKey() => _nextGroupKey++;

    /// <summary>
    /// Group existing free-floating pictures into a new group shape. The group's bounding box is the
    /// union of the members' sheet rectangles; the members keep their on-sheet positions and sizes
    /// (the group uses an identity child coordinate space). The group is built on the next save.
    /// (A first-class group API is added in a later phase.)
    /// </summary>
    public IXLPictureGroup Group(params IXLPicture[] pictures)
    {
        ArgumentNullException.ThrowIfNull(pictures);

        var members = CollectDistinctMembers(pictures);
        if (members.Count < 2)
            throw new ArgumentException("A group needs at least two distinct pictures.", nameof(pictures));

        var (minX, minY, maxX, maxY) = ValidateMembersAndComputeBounds(members);

        var groupKey = AllocateGroupKey();
        AssignIdentityGroupInfo(members, groupKey);

        // members.Count >= 2 is guaranteed above, so the min/max sentinels are always
        // overwritten in the loop, and minX <= maxX / minY <= maxY by construction (each
        // member feeds x to minX and x+width to maxX). The subtraction cannot underflow.
#pragma warning disable S3949 // Calculations should not overflow
        PendingGroups.Add(new XLPendingGroup([.. members], minX, minY, maxX - minX, maxY - minY));
#pragma warning restore S3949
        return new XLPictureGroupView(_worksheet, groupKey);
    }

    private static List<XLPicture> CollectDistinctMembers(IXLPicture[] pictures)
    {
        var members = new List<XLPicture>(pictures.Length);
        foreach (var picture in pictures)
        {
            if (picture is null)
                throw new ArgumentException("The pictures cannot contain a null element.", nameof(pictures));
            if (picture is not XLPicture member)
                throw new ArgumentException("The picture was not created by this library.", nameof(pictures));
            if (!members.Contains(member))
                members.Add(member);
        }

        return members;
    }

    private (long MinX, long MinY, long MaxX, long MaxY) ValidateMembersAndComputeBounds(List<XLPicture> pictures)
    {
        var wb = _worksheet.Workbook;
        long minX = long.MaxValue, minY = long.MaxValue, maxX = long.MinValue, maxY = long.MinValue;

        foreach (var member in pictures)
        {
            if (!ReferenceEquals(member.Worksheet, _worksheet))
                throw new ArgumentException("All pictures must belong to this worksheet.", nameof(pictures));
            if (member.IsInGroup)
                throw new ArgumentException($"Picture '{member.Name}' is already in a group.", nameof(pictures));
            if (member.Placement != XLPicturePlacement.FreeFloating)
                throw new NotSupportedException(
                    $"Only free-floating pictures can be grouped; '{member.Name}' is '{member.Placement}'. Call MoveTo(left, top) first.");
            if (string.IsNullOrEmpty(member.RelId))
                throw new NotSupportedException($"Picture '{member.Name}' must be saved before it can be grouped.");

            var x = ToEmu(member.Left, wb.DpiX);
            var y = ToEmu(member.Top, wb.DpiY);
            minX = Math.Min(minX, x);
            minY = Math.Min(minY, y);
            maxX = Math.Max(maxX, x + ToEmu(member.Width, wb.DpiX));
            maxY = Math.Max(maxY, y + ToEmu(member.Height, wb.DpiY));
        }

        return (minX, minY, maxX, maxY);
    }

    private static void AssignIdentityGroupInfo(List<XLPicture> members, long groupKey)
    {
        foreach (var member in members)
        {
            // Identity child space (chOff/chExt == off/ext), so a child's coordinates equal its
            // sheet coordinates: composed scale 1, composed offset 0.
            member.GroupInfo = new XLPictureGroup
            {
                ScaleX = 1.0,
                ScaleY = 1.0,
                OffsetX = 0.0,
                OffsetY = 0.0,
                GroupKey = groupKey,
                LoadedWidthPx = member.Width,
                LoadedHeightPx = member.Height,
                LoadedLeftPx = member.Left,
                LoadedTopPx = member.Top,
            };
        }
    }

    private static long ToEmu(int pixels, double dpi) => Convert.ToInt64(914400L * pixels / dpi);

    /// <summary>
    /// Loads the pictures and shapes of a drawing through <paramref name="loadPictures"/>, and keeps
    /// the names of the drawing's shapes that loaded as neither (issue #796): a linked picture, a
    /// picture under a wrapper the loader does not read, a connector, a chart. The save keeps those
    /// as the file has them, so no picture is given one of their names.
    /// </summary>
    /// <param name="shapeNames">The name of each shape of the drawing, once per shape.</param>
    /// <param name="loadPictures">Adds the drawing's pictures to this collection.</param>
    /// <remarks>
    /// While the pictures load, every name in the drawing is taken, so a picture the file leaves
    /// unnamed is not named after a shape further on. Each loaded picture then accounts for one use
    /// of its name, so a name two shapes share is still kept when only one of them loaded.
    /// </remarks>
    internal void LoadDrawing(IEnumerable<string> shapeNames, Action loadPictures)
    {
        var uses = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        foreach (var name in shapeNames)
            uses[name] = uses.GetValueOrDefault(name) + 1;

        _keptShapeNames = new HashSet<string>(uses.Keys, StringComparer.OrdinalIgnoreCase);
        loadPictures();

        // A shape that loaded into the worksheet's shapes holds its name there, for as long as it
        // keeps it (spec 15), so it is not kept here.
        var loadedNames = _pictures.Select(picture => picture.Name)
            .Concat(((IEnumerable<XLShape>)(XLShapes)_worksheet.Shapes).Select(shape => shape.Name));
        foreach (var name in loadedNames)
        {
            if (uses.TryGetValue(name, out var count))
                uses[name] = count - 1;
        }

        _keptShapeNames = new HashSet<string>(uses.Where(use => use.Value > 0).Select(use => use.Key),
            StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Whether <paramref name="name"/>, in any case, is the name of a shape the drawing keeps, of a
    /// shape of the worksheet, or of a picture of this collection other than <paramref name="changing"/>.
    /// </summary>
    /// <param name="name">The name to check.</param>
    /// <param name="changing">The picture that is to take the name, or <c>null</c> for a new one.</param>
    internal bool IsNameTaken(string name, XLPicture? changing = null) =>
        _keptShapeNames.Contains(name) ||
        ((XLShapes)_worksheet.Shapes).IsNameUsed(name) ||
        _nameUses.GetValueOrDefault(name) >
        (changing is { IsInCollection: true } && string.Equals(changing.Name, name, StringComparison.OrdinalIgnoreCase) ? 1 : 0);

    /// <summary>Whether a picture of this collection other than <paramref name="changing"/> has the id <paramref name="id"/>.</summary>
    /// <param name="id">The id to check.</param>
    /// <param name="changing">The picture that is to take the id.</param>
    internal bool IsIdTaken(int id, XLPicture changing) =>
        _idUses.GetValueOrDefault(id) > (changing.IsInCollection && changing.Id == id ? 1 : 0);

    /// <summary>Whether a picture of this collection has the id <paramref name="id"/>.</summary>
    internal bool IsIdUsed(int id) => _idUses.ContainsKey(id);

    /// <summary>
    /// The id a new picture takes: one more than the highest id of any picture in
    /// <paramref name="workbook"/>, or 1 when it has none.
    /// </summary>
    internal static int NextFreeId(XLWorkbook workbook)
    {
        var highest = 0;
        foreach (var worksheet in workbook.Worksheets)
        {
            if (((XLPictures)worksheet.Pictures).HighestId is { } id)
                highest = Math.Max(highest, id);
        }

        return highest + 1;
    }

    /// <summary>The highest id of a picture of this collection, or <c>null</c> when it is empty.</summary>
    private int? HighestId => _idUses.Count == 0 ? null : _highestId ??= ReadHighestId();

    /// <summary>How many times the highest id was found by reading every id. A probe for tests.</summary>
    internal int HighestIdReads { get; private set; }

    /// <summary>Reads every id for the highest, and the highest below it. The collection is not empty.</summary>
    private int ReadHighestId()
    {
        HighestIdReads++;
        int? highest = null;
        int? below = null;
        foreach (var id in _idUses.Keys)
        {
            if (highest is not { } h || id > h)
            {
                below = highest;
                highest = id;
            }
            else if (below is not { } b || id > b)
            {
                below = id;
            }
        }

        _belowHighestId = below;
        return highest.GetValueOrDefault();
    }

    /// <summary>Called by <paramref name="picture"/> when its name changes from <paramref name="oldName"/>.</summary>
    internal void OnRenamed(XLPicture picture, string oldName)
    {
        Release(_nameUses, oldName);
        TakeName(picture.Name);
    }

    /// <summary>Called by <paramref name="picture"/> when its id changes from <paramref name="oldId"/>.</summary>
    internal void OnIdChanged(XLPicture picture, int oldId)
    {
        ReleaseId(oldId);
        TakeId(picture.Id);
    }

    private void Attach(XLPicture picture)
    {
        _pictures.Add(picture);
        picture.IsInCollection = true;
        TakeName(picture.Name);
        TakeId(picture.Id);
    }

    private void Detach(XLPicture picture)
    {
        _pictures.Remove(picture);
        picture.IsInCollection = false;
        Release(_nameUses, picture.Name);
        ReleaseId(picture.Id);
    }

    /// <summary>A picture not yet named, being created, has the empty name, which takes nothing.</summary>
    private void TakeName(string name)
    {
        if (name.Length != 0)
            Take(_nameUses, name);
    }

    private void TakeId(int id)
    {
        if (!Take(_idUses, id))
            return;

        if (_idUses.Count == 1)
        {
            _highestId = id;
            _belowHighestId = null;
            return;
        }

        // A highest not known is read when it is next asked for.
        if (_highestId is not { } highest)
            return;

        if (id > highest)
        {
            _belowHighestId = highest;
            _highestId = id;
        }
        else if (_belowHighestId < id || _idUses.Count == 2)
        {
            // With two ids, the one below the highest is this one.
            _belowHighestId = id;
        }
    }

    private void ReleaseId(int id)
    {
        if (!Release(_idUses, id))
            return;

        if (id == _highestId)
        {
            _highestId = _belowHighestId;
            _belowHighestId = null;
        }
        else if (id == _belowHighestId)
        {
            _belowHighestId = null;
        }
    }

    /// <summary>Takes one use of <paramref name="key"/>; returns whether it is the first.</summary>
    private static bool Take<TKey>(Dictionary<TKey, int> uses, TKey key) where TKey : notnull
    {
        var count = uses.GetValueOrDefault(key);
        uses[key] = count + 1;
        return count == 0;
    }

    /// <summary>Releases one use of <paramref name="key"/>; returns whether no use of it is left.</summary>
    private static bool Release<TKey>(Dictionary<TKey, int> uses, TKey key) where TKey : notnull
    {
        if (!uses.TryGetValue(key, out var count))
            return false;

        if (count > 1)
        {
            uses[key] = count - 1;
            return false;
        }

        uses.Remove(key);
        return true;
    }

    /// <summary>
    /// Throws when <paramref name="name"/> is taken (see <see cref="IsNameTaken"/>). A named <c>Add</c>
    /// checks before it creates the picture, so a taken name leaves the collection as it was (#806).
    /// </summary>
    internal void ThrowIfNameTaken(string name, XLPicture? changing = null)
    {
        if (IsNameTaken(name, changing))
            throw new ArgumentException($"The picture name '{name}' already exists.");
    }

    private string GetNextPictureName()
    {
        var pictureNumber = Count;
        while (IsNameTaken($"Picture {pictureNumber}"))
        {
            pictureNumber++;
        }

        return $"Picture {pictureNumber}";
    }
}

/// <summary>
/// A group to be created on the next save: its member pictures and the group's bounding box in EMU
/// (used for the group's <c>off</c>/<c>ext</c> and <c>chOff</c>/<c>chExt</c>).
/// </summary>
internal sealed record XLPendingGroup(
    IReadOnlyList<XLPicture> Members,
    long OffsetX,
    long OffsetY,
    long ExtentCx,
    long ExtentCy);
