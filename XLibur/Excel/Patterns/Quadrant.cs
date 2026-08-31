using System;
using System.Collections.Generic;
using System.Linq;
using XLibur.Excel.Coordinates;
using XLibur.Excel.Ranges;

namespace XLibur.Excel.Patterns;

/// <summary>
/// Implementation of QuadTree adapted to Excel worksheet specifics. Differences with the classic implementation
/// are that the topmost level is split to 128 square parts (2 columns of 64 blocks, each 8192*8192 cells) and that splitting
/// the quadrant onto 4 smaller quadrants does not depend on the number of items in this quadrant. When the range is added to the
/// QuadTree it is placed on the bottommost level where it fits to a single quadrant. That means, row-wide and column-wide ranges
/// are always placed at the level 0, and the smaller the range is the deeper it goes down the tree. This approach eliminates
/// the need of transferring ranges between levels.
/// </summary>
internal class Quadrant
{
    #region Public Properties

    /// <summary>
    /// Smaller quadrants which the current one is split to. Is NULL until ranges are added to child quadrants.
    /// </summary>
    public IReadOnlyList<Quadrant>? Children { get; private set; }

    /// <summary>
    /// The level of current quadrant. Top most has level 0, child quadrants has levels (Level + 1).
    /// </summary>
    public byte Level { get; }

    /// <summary>
    /// Minimum column included in this quadrant.
    /// </summary>
    public int MinimumColumn { get; }

    /// <summary>
    /// Minimum row included in this quadrant.
    /// </summary>
    public int MinimumRow { get; }

    /// <summary>
    /// Maximum column included in this quadrant.
    /// </summary>
    public int MaximumColumn { get; }

    /// <summary>
    /// Maximum row included in this quadrant.
    /// </summary>
    public int MaximumRow { get; }

    /// <summary>
    /// Collection of ranges belonging to this quadrant (does not include ranges from child quadrants).
    /// </summary>
    public IEnumerable<IXLAddressable>? Ranges
    {
        get => _ranges?.Values.AsEnumerable();
    }

    /// <summary>
    /// The number of current quadrant by horizontal axis.
    /// </summary>
    public short X { get; private set; }

    /// <summary>
    /// The number of current quadrant by vertical axis.
    /// </summary>
    public short Y { get; private set; }

    #endregion Public Properties

    #region Constructors

    public Quadrant() : this(0, 0, 0)
    { }

    private Quadrant(byte level, short x, short y)
    {
        Level = level;
        X = x;
        Y = y;

        MinimumColumn = (Level == 0) ? 1 : 1 + XLHelper.MaxColumnNumber / (int)Math.Pow(2, Level) * X;
        MinimumRow = (Level == 0) ? 1 : 1 + XLHelper.MaxColumnNumber / (int)Math.Pow(2, Level) * Y; //MaxColumnNumber here is not a mistake
        MaximumColumn = (Level == 0)
            ? XLHelper.MaxColumnNumber
            : XLHelper.MaxColumnNumber / (int)Math.Pow(2, Level) * (X + 1);
        MaximumRow = (Level == 0)
            ? XLHelper.MaxRowNumber
            : XLHelper.MaxColumnNumber / (int)Math.Pow(2, Level) * (Y + 1); //MaxColumnNumber here is not a mistake
    }

    #endregion Constructors

    #region Public Methods

    /// <summary>
    /// Add a range to the quadrant or to one of the child quadrants (recursively).
    /// </summary>
    /// <returns>True, if range was successfully added, false if it has been added before.</returns>
    public bool Add(IXLAddressable range)
    {
        return Add(range, Area.FromRangeAddress(range.RangeAddress));
    }

    private bool Add(IXLAddressable range, in Area area)
    {
        var res = false;
        var children = Children ?? CreateChildren().ToList();
        var addToChild = false;
        foreach (var childQuadrant in children)
        {
            if (childQuadrant.Covers(in area))
            {
                res |= childQuadrant.Add(range, in area);
                addToChild = true;
                break;
            }
        }

        if (!addToChild)
            res = AddInternal(range);

        if (Children == null && addToChild)
            Children = children;

        if (res)
            _subtreeCount++;

        return res;
    }

    /// <summary>
    /// Get all ranges from the quadrant and all child quadrants (recursively).
    /// </summary>
    public IEnumerable<IXLAddressable> GetAll()
    {
        if (_subtreeCount == 0)
            yield break;

        if (Ranges != null)
        {
            foreach (var range in Ranges)
                yield return range;
        }

        if (Children != null)
        {
            foreach (var childQuadrant in Children)
            {
                if (childQuadrant._subtreeCount == 0)
                    continue;

                var childRanges = childQuadrant.GetAll();
                foreach (var range in childRanges)
                    yield return range;
            }
        }
    }

    /// <summary>
    /// Get all ranges from the quadrant and all child quadrants (recursively) that intersect the specified address.
    /// </summary>
    public IEnumerable<IXLAddressable> GetIntersectedRanges(IXLRangeAddress rangeAddress)
    {
        return GetIntersectedRanges(Area.FromRangeAddress(rangeAddress));
    }

    /// <summary>
    /// Same traversal as <see cref="GetIntersectedRanges(IXLRangeAddress)"/>, but keyed on the
    /// normalised rectangle rather than the address: <see cref="XLRangeAddress.Intersects(IXLRangeAddress)"/>
    /// assumes both sides are already normalised and gives wrong answers for a reversed address,
    /// which is exactly the input this index cannot assume once it holds more than a handful of
    /// ranges (see <see cref="Quadrant{T}"/> and its promotion threshold).
    /// </summary>
    private IEnumerable<IXLAddressable> GetIntersectedRanges(Area area)
    {
        if (_subtreeCount == 0)
            yield break;

        if (Ranges != null)
        {
            foreach (var range in Ranges)
            {
                if (Area.FromRangeAddress(range.RangeAddress).Intersects(area))
                    yield return range;
            }
        }

        foreach (var range in GetIntersectedRangesFromChildren(area))
            yield return range;
    }

    /// <summary>
    /// Get all ranges from the quadrant and all child quadrants (recursively) that cover the specified address.
    /// </summary>
    public IEnumerable<IXLAddressable> GetIntersectedRanges(IXLAddress address)
    {
        if (_subtreeCount == 0)
            yield break;

        if (Ranges != null)
        {
            // Reads through XLAddressableHelper.Contains (Area-based) rather than
            // IXLRangeAddress.Contains directly: the latter assumes FirstAddress is the top-left
            // corner, which does not hold for a range added with its corners reversed - see
            // ReversedRangeGeometryTests.MergedReversedRangeIsRecognisedBeforeAndAfterPromotion.
            var xlAddress = (XLAddress)address;
            foreach (var range in Ranges)
            {
                if (XLAddressableHelper.Contains(range, in xlAddress))
                    yield return range;
            }
        }

        foreach (var range in GetIntersectedRangesFromChildren(address))
            yield return range;
    }

    /// <summary>
    /// Whether any range in this quadrant or its children covers the address. Same traversal as
    /// <see cref="GetIntersectedRanges(IXLAddress)"/>, but as a plain recursion so that answering
    /// a yes/no question does not allocate an iterator per level. Used by the merged-range test
    /// that runs on every cell write.
    /// </summary>
    public bool CoversAnyRange(in XLAddress address)
    {
        if (_subtreeCount == 0)
            return false;

        if (_ranges is not null)
        {
            foreach (var range in _ranges.Values)
            {
                if (XLAddressableHelper.Contains(range, in address))
                    return true;
            }
        }

        var children = Children;
        if (children is null)
            return false;

        // Indexed rather than foreach: Children is typed as IReadOnlyList<Quadrant>, so a foreach
        // would allocate an interface enumerator at every level of the recursion — which is the
        // one thing this method exists to avoid.
        for (var i = 0; i < children.Count; i++)
        {
            var childQuadrant = children[i];
            if (childQuadrant.Covers(in address) && childQuadrant.CoversAnyRange(in address))
                return true;
        }

        return false;
    }

    private IEnumerable<IXLAddressable> GetIntersectedRangesFromChildren(Area area)
    {
        if (Children == null)
            yield break;

        foreach (var childQuadrant in Children)
        {
            if (childQuadrant._subtreeCount > 0 && childQuadrant.Intersects(in area))
            {
                foreach (var range in childQuadrant.GetIntersectedRanges(area))
                    yield return range;
            }
        }
    }

    private IEnumerable<IXLAddressable> GetIntersectedRangesFromChildren(IXLAddress address)
    {
        if (Children == null)
            yield break;

        foreach (var childQuadrant in Children)
        {
            if (childQuadrant._subtreeCount > 0 && childQuadrant.Covers(in address))
            {
                foreach (var range in childQuadrant.GetIntersectedRanges(address))
                    yield return range;
            }
        }
    }

    /// <summary>
    /// Remove the range from the quadrant or from child quadrants (recursively).
    /// </summary>
    /// <returns>True if the range was removed, false if it does not exist in the QuadTree.</returns>
    public bool Remove(IXLRangeAddress rangeAddress)
    {
        return Remove(rangeAddress, Area.FromRangeAddress(rangeAddress));
    }

    private bool Remove(IXLRangeAddress rangeAddress, in Area area)
    {
        if (_subtreeCount == 0)
            return false;

        var res = false;

        var coveredByChild = false;
        if (Children != null)
        {
            foreach (var childQuadrant in Children)
            {
                if (childQuadrant.Covers(in area))
                {
                    res |= childQuadrant.Remove(rangeAddress, in area);
                    coveredByChild = true;
                }
            }
        }

        if (!coveredByChild && _ranges?.Remove(rangeAddress) == true)
            res = true;

        if (res)
            _subtreeCount--;

        return res;
    }

    /// <summary>
    /// Remove all the ranges matching specified criteria from the quadrant and its child quadrants (recursively).
    /// Don't use it for searching intersections as it would be much less efficient than <see cref="GetIntersectedRanges(IXLRangeAddress)"/>.
    /// </summary>
    /// <remarks>
    /// Eager rather than lazy: every caller wants the whole set (or just its size), and removal has to
    /// stay all-or-nothing for <see cref="_subtreeCount"/> to keep matching the contents — a half-consumed
    /// lazy walk used to yield ranges it had not removed yet.
    /// </remarks>
    public IReadOnlyList<IXLAddressable> RemoveAll(Predicate<IXLAddressable> predicate)
    {
        var removed = new List<IXLAddressable>();
        RemoveAllInto(predicate, removed);
        return removed;
    }

    /// <summary>
    /// Recursive worker for <see cref="RemoveAll"/>. Returns the number of ranges removed from this
    /// quadrant and its descendants, so each level can adjust its own <see cref="_subtreeCount"/>.
    /// </summary>
    private int RemoveAllInto(Predicate<IXLAddressable> predicate, List<IXLAddressable> removed)
    {
        // Child quadrants are created on demand but never destroyed, so a long-lived index that has
        // seen many add/remove cycles keeps a skeleton of empty quadrants. Without this test every
        // removal would walk that skeleton, which is what made a Clear/CopyTo loop quadratic (#271).
        if (_subtreeCount == 0)
            return 0;

        var count = RemoveOwnRanges(predicate, removed);

        var children = Children;
        if (children != null)
        {
            for (var i = 0; i < children.Count; i++)
                count += children[i].RemoveAllInto(predicate, removed);
        }

        _subtreeCount -= count;
        return count;
    }

    /// <summary>
    /// Removes this quadrant's own matching ranges, appending them to <paramref name="removed"/>, and
    /// returns how many went. The matches are collected before any are removed because the dictionary
    /// cannot be modified while its values are being enumerated.
    /// </summary>
    private int RemoveOwnRanges(Predicate<IXLAddressable> predicate, List<IXLAddressable> removed)
    {
        if (_ranges == null)
            return 0;

        List<IXLRangeAddress>? keysToRemove = null;
        foreach (var range in _ranges.Values)
        {
            if (!predicate(range))
                continue;

            (keysToRemove ??= new List<IXLRangeAddress>()).Add(range.RangeAddress);
            removed.Add(range);
        }

        if (keysToRemove == null)
            return 0;

        var count = 0;
        foreach (var keyToRemove in keysToRemove)
        {
            if (_ranges.Remove(keyToRemove))
                count++;
        }

        return count;
    }

    #endregion Public Methods

    #region Private Fields

    /// <summary>
    /// Maximum depth of the QuadTree. Value 10 corresponds to the smallest quadrants having size 16*16 cells.
    /// </summary>
    private const byte MAX_LEVEL = 10;

    /// <summary>
    /// Collection of ranges belonging to the current quadrant (that cannot fit into child quadrants).
    /// </summary>
    private Dictionary<IXLRangeAddress, IXLAddressable>? _ranges;

    /// <summary>
    /// Number of ranges held by this quadrant and every descendant. <see cref="Children"/> is created
    /// lazily but never torn down, so an index that has seen many add/remove cycles accumulates empty
    /// quadrants; this count lets every traversal skip a subtree that has nothing in it instead of
    /// walking that skeleton.
    /// </summary>
    private int _subtreeCount;

    #endregion Private Fields

    #region Private Methods

    /// <summary>
    /// Add a range to the collection of quadrant's own ranges.
    /// </summary>
    /// <returns>True if the range was successfully added, false if it had been added before.</returns>
    private bool AddInternal(IXLAddressable range)
    {
        _ranges ??= new Dictionary<IXLRangeAddress, IXLAddressable>();
        return _ranges.TryAdd(range.RangeAddress, range);
    }

    /// <summary>
    /// Check if the current quadrant fully covers the specified rectangle.
    /// </summary>
    private bool Covers(in Area area)
    {
        return MinimumColumn <= area.LeftColumn &&
               MaximumColumn >= area.RightColumn &&
               MinimumRow <= area.TopRow &&
               MaximumRow >= area.BottomRow;
    }

    /// <summary>
    /// Check if the current quadrant covers the specified address. Overload taking the concrete
    /// struct, so <see cref="CoversAnyRange"/> does not box on every level of the recursion.
    /// </summary>
    private bool Covers(in XLAddress address)
    {
        return MinimumColumn <= address.ColumnNumber &&
               MaximumColumn >= address.ColumnNumber &&
               MinimumRow <= address.RowNumber &&
               MaximumRow >= address.RowNumber;
    }

    /// <summary>
    /// Check if the current quadrant covers the specified address.
    /// </summary>
    private bool Covers(in IXLAddress address)
    {
        return MinimumColumn <= address.ColumnNumber &&
               MaximumColumn >= address.ColumnNumber &&
               MinimumRow <= address.RowNumber &&
               MaximumRow >= address.RowNumber;
    }

    /// <summary>
    /// Check if the current quadrant intersects the specified rectangle.
    /// </summary>
    private bool Intersects(in Area area)
    {
        return ((MinimumRow <= area.TopRow && area.TopRow <= MaximumRow) ||
                (area.TopRow <= MinimumRow && MinimumRow <= area.BottomRow))
               &&
               ((MinimumColumn <= area.LeftColumn && area.LeftColumn <= MaximumColumn) ||
                (area.LeftColumn <= MinimumColumn && MinimumColumn <= area.RightColumn));
    }

    /// <summary>
    /// Create a collection of child quadrants dividing the current one.
    /// </summary>
    private IEnumerable<Quadrant> CreateChildren()
    {
        var childLevel = (byte)(Level + 1);
        if (childLevel > MAX_LEVEL)
            yield break;
        byte xCount = 2; // Always divide in halves
        var yCount = (byte)(Level == 0 ? (XLHelper.MaxRowNumber / XLHelper.MaxColumnNumber) : 2); // Level 0 divide onto 64 parts, the rest - on halves

        for (byte dy = 0; dy < yCount; dy++)
        {
            for (byte dx = 0; dx < xCount; dx++)
            {
                yield return new Quadrant(childLevel, (short)(X * 2 + dx), (short)(Y * 2 + dy));
            }
        }
    }

    #endregion Private Methods
}

/// <summary>
/// A generic version of <see cref="Quadrant"/>
/// </summary>
internal sealed class Quadrant<T> : Quadrant
    where T : IXLAddressable
{
    public new IEnumerable<T>? Ranges => base.Ranges?.Cast<T>();

    public bool Add(T range)
    {
        return base.Add(range);
    }

    public new IEnumerable<T> GetAll()
    {
        return base.GetAll().Cast<T>();
    }

    public new IEnumerable<T> GetIntersectedRanges(IXLRangeAddress rangeAddress)
    {
        return base.GetIntersectedRanges(rangeAddress).Cast<T>();
    }

    public new IEnumerable<T> GetIntersectedRanges(IXLAddress address)
    {
        return base.GetIntersectedRanges(address).Cast<T>();
    }

    public bool Remove(T range)
    {
        return Remove(range.RangeAddress);
    }
    public IEnumerable<T> RemoveAll(Predicate<T> predicate)
    {
        return base.RemoveAll(r => predicate((T)r)).Cast<T>();
    }
}
