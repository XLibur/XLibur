using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Numerics;
using System.Runtime.InteropServices;
using XLibur.Extensions;

namespace XLibur.Excel;

internal sealed partial class Slice<TElement>
{
    /// <summary>
    /// <para>
    /// Memory efficient look up table. The table is 2-level structure,
    /// where elements of the the top level are potentially nullable
    /// references to buckets of up-to 32 items in bottom level.
    /// </para>
    /// <para>
    /// Both level can increase size through doubling, through
    /// only the top one can be indefinite size.
    /// </para>
    /// </summary>
    private sealed class Lut<T>
    {
        private const int BottomLutBits = 5;
        private const int BottomLutMask = (1 << BottomLutBits) - 1;

        /// <summary>
        /// The default value lut ref returns for elements not defined in the lut.
        /// </summary>
        private static readonly T DefaultValue = default!;

        /// <summary>
        /// A sparse array of values in the lut. The top level always allocated at least one element.
        /// </summary>
        private LutBucket[] _buckets = new LutBucket[1];

        /// <summary>
        /// Get maximal node that is used. Return -1 if LUT is unused.
        /// </summary>
        internal int MaxUsedIndex { get; private set; } = -1;

        /// <summary>
        /// Does LUT contains at least one used element?
        /// </summary>
        internal bool IsEmpty => MaxUsedIndex < 0;

        /// <summary>
        /// Get a value at specified index.
        /// </summary>
        /// <param name="index">Index, starting at 0.</param>
        /// <returns>Reference to an element at index, if the element is used, otherwise <see cref="DefaultValue"/>.</returns>
        internal ref readonly T Get(int index)
        {
            var (topIdx, bottomIdx) = SplitIndex(index);
            if (topIdx >= _buckets.Length)
                return ref DefaultValue;

            if (!IsUsed(topIdx, bottomIdx))
                return ref DefaultValue;

            var nodes = _buckets[topIdx].Nodes!;
            return ref nodes[bottomIdx];
        }

        /// <summary>
        /// Does the index set a mask of used index (=was value set and not cleared)?
        /// </summary>
        internal bool IsUsed(int index)
        {
            var (topIdx, bottomIdx) = SplitIndex(index);
            if (topIdx >= _buckets.Length)
                return false;

            return IsUsed(topIdx, bottomIdx);
        }

        /// <summary>
        /// Set/clar an element at index to a specified value.
        /// The used flag will be if the value is <c>default</c> or not.
        /// </summary>
        internal void Set(int index, T value)
        {
            var (topIdx, bottomIdx) = SplitIndex(index);

            SetValue(value, topIdx, bottomIdx);

            var valueIsDefault = EqualityComparer<T>.Default.Equals(value, DefaultValue);
            if (valueIsDefault)
                ClearBitmap(topIdx, bottomIdx);
            else
                SetBitmap(topIdx, bottomIdx);

            if (_buckets[topIdx].Bitmap == 0)
                _buckets[topIdx] = new LutBucket(null, 0);

            RecalculateMaxIndex(index);
        }

        /// <summary>
        /// Fast path for setting a value that the caller guarantees is not <c>default</c>.
        /// Skips the <see cref="EqualityComparer{T}"/> check and always sets the bitmap bit.
        /// </summary>
        internal void SetNonDefault(int index, T value)
        {
            var (topIdx, bottomIdx) = SplitIndex(index);
            SetValue(value, topIdx, bottomIdx);
            SetBitmap(topIdx, bottomIdx);

            if (index > MaxUsedIndex)
                MaxUsedIndex = index;
        }

        private void SetValue(T value, int topIdx, int bottomIdx)
        {
            var topSize = _buckets.Length;
            if (topIdx >= topSize)
            {
                do
                {
                    topSize *= 2;
                } while (topIdx >= topSize);

                Array.Resize(ref _buckets, topSize);
            }

            var bucket = _buckets[topIdx];
            var bottomBucketExists = bucket.Nodes is not null;
            if (!bottomBucketExists)
            {
                var initialSize = 4;
                while (bottomIdx >= initialSize)
                    initialSize *= 2;

                _buckets[topIdx] = bucket = new LutBucket(new T[initialSize], 0);
            }
            else
            {
                // Bottom exists, but might not be large enough
                var bottomSize = bucket.Nodes!.Length;
                if (bottomIdx >= bottomSize)
                {
                    do
                    {
                        bottomSize *= 2;
                    } while (bottomIdx >= bottomSize);

                    var bucketNodes = bucket.Nodes;
                    Array.Resize(ref bucketNodes, bottomSize);
                    _buckets[topIdx] = bucket = new LutBucket(bucketNodes, bucket.Bitmap);
                }
            }

            bucket.Nodes![bottomIdx] = value;
        }

        private static (int TopLevelIndex, int BottomLevelIndex) SplitIndex(int index)
        {
            var topIdx = index >> BottomLutBits;
            var bottomIdx = index & BottomLutMask;
            return (topIdx, bottomIdx);
        }

        private bool IsUsed(int topIdx, int bottomIdx)
            => (_buckets[topIdx].Bitmap & (1 << bottomIdx)) != 0;

        private void SetBitmap(int topIdx, int bottomIdx)
            => _buckets[topIdx] = new LutBucket(_buckets[topIdx].Nodes, _buckets[topIdx].Bitmap | (uint)1 << bottomIdx);

        private void ClearBitmap(int topIdx, int bottomIdx)
            => _buckets[topIdx] = new LutBucket(_buckets[topIdx].Nodes, _buckets[topIdx].Bitmap & ~((uint)1 << bottomIdx));

        private void RecalculateMaxIndex(int index)
        {
            if (MaxUsedIndex <= index)
                MaxUsedIndex = CalculateMaxIndex();
        }

        private int CalculateMaxIndex()
        {
            for (var bucketIdx = _buckets.Length - 1; bucketIdx >= 0; --bucketIdx)
            {
                var bitmap = _buckets[bucketIdx].Bitmap;
                if (bitmap != 0)
                {
                    return (bucketIdx << BottomLutBits) + bitmap.GetHighestSetBit();
                }
            }

            return -1;
        }

        /// <summary>
        /// A bucket of bottom layer of LUT. Each bucket has up-to 32 elements.
        /// </summary>
        [StructLayout(LayoutKind.Sequential, Pack = 4)]
        private readonly struct LutBucket
        {
            public readonly T[]? Nodes;

            /// <summary>
            /// <para>
            /// A bitmap array that indicates which nodes have a set/no-default values values
            /// (1 = value has been set and there is an element in the <see cref="_buckets"/>,
            /// 0 = value hasn't been set and <see cref="_buckets"/> might exist or not).
            /// If the element at some index is not is not set and lut is asked for a value,
            /// it should return <see cref="DefaultValue"/>.
            /// </para>
            /// <para>
            /// The length of the bitmap array is same as the <see cref="_buckets"/>, for each
            /// bottom level bucket, the element of index 0 in the bucket is represented by
            /// lowest bit, element 31 is represented by highest bit.
            /// </para>
            /// <para>
            /// This is useful to make a distinction between a node that is empty
            /// and a node that had it's value se to <see cref="DefaultValue"/>.
            /// </para>
            /// </summary>
            public readonly uint Bitmap;

            internal LutBucket(T[]? nodes, uint bitmap)
            {
                Nodes = nodes;
                Bitmap = bitmap;
            }
        }

        /// <summary>
        /// Enumerator of LUT used values from low index to high.
        /// </summary>
        internal struct LutEnumerator
        {
            private readonly Lut<T> _lut;
            private readonly int _endIdx;
            private int _idx;

            /// <summary>
            /// Create a new enumerator from subset of elements.
            /// </summary>
            /// <param name="lut">Lookup table to traverse.</param>
            /// <param name="startIdx">First desired index, included.</param>
            /// <param name="endIdx">Last desired index, included.</param>
            internal LutEnumerator(Lut<T> lut, int startIdx, int endIdx)
            {
                Debug.Assert(startIdx <= endIdx);
                _lut = lut;
                _idx = startIdx - 1;
                _endIdx = endIdx;
            }

            public ref T Current => ref _lut._buckets[_idx >> BottomLutBits].Nodes![_idx & BottomLutMask];

            /// <summary>
            /// Index of current element in the LUT. Only valid, if enumerator is valid.
            /// </summary>
            public int Index => _idx;

            public bool MoveNext()
            {
                var usedIndex = GetNextUsedIndexAtOrLater(_idx + 1);
                if (usedIndex > _endIdx)
                    return false;

                _idx = usedIndex;
                return true;
            }

            private int GetNextUsedIndexAtOrLater(int index)
            {
                var buckets = _lut._buckets;
                var (topIdx, bottomIdx) = SplitIndex(index);

                while (topIdx < buckets.Length)
                {
                    var setBitIndex = buckets[topIdx].Bitmap.GetLowestSetBitAbove(bottomIdx);
                    if (setBitIndex >= 0)
                        return topIdx * 32 + setBitIndex;

                    ++topIdx;
                    bottomIdx = 0;
                }

                // We are the end of LUT
                return int.MaxValue;
            }
        }

        /// <summary>
        /// Enumerator of LUT used values from high index to low index.
        /// </summary>
        internal struct ReverseLutEnumerator
        {
            private readonly Lut<T> _lut;
            private readonly int _startIdx;
            private int _idx;

            internal ReverseLutEnumerator(Lut<T> lut, int startIdx, int endIdx)
            {
                Debug.Assert(startIdx <= endIdx);
                _lut = lut;
                _idx = endIdx + 1;
                _startIdx = startIdx;
            }

            public ref T Current => ref _lut._buckets[_idx >> BottomLutBits].Nodes![_idx & BottomLutMask];

            public int Index => _idx;

            public bool MoveNext()
            {
                var usedIndex = GetPrevIndexAtOrBefore(_idx - 1);
                if (usedIndex < _startIdx)
                    return false;

                _idx = usedIndex;
                return true;
            }

            private int GetPrevIndexAtOrBefore(int index)
            {
                var buckets = _lut._buckets;
                var (topIdx, bottomIdx) = SplitIndex(index);
                if (topIdx >= buckets.Length)
                {
                    topIdx = buckets.Length - 1;
                    bottomIdx = 31;
                }

                while (topIdx >= 0)
                {
                    var setBitIndex = buckets[topIdx].Bitmap.GetHighestSetBitBelow(bottomIdx);
                    if (setBitIndex >= 0)
                        return topIdx * 32 + setBitIndex;

                    --topIdx;
                    bottomIdx = 31;
                }

                return int.MinValue;
            }
        }
    }

    /// <summary>
    /// Compact per-row storage. For narrow rows (all columns &lt; 32), stores a flat
    /// array indexed by column with a 32-bit bitmap. For wide rows, delegates to
    /// a full <see cref="Lut{T}"/>. Default struct represents an empty row, allowing
    /// storage in the outer <see cref="Lut{T}"/> without per-row heap allocations.
    /// </summary>
    private struct RowData : IEquatable<RowData>
    {
        private static readonly TElement ElementDefault = default!;

        /// <summary>
        /// Discriminated storage: null (empty), TElement[] (narrow), or Lut&lt;TElement&gt; (wide).
        /// </summary>
        private object? _storage;

        /// <summary>
        /// In narrow mode, bit i set means column i is used. Unused in wide mode.
        /// </summary>
        private uint _bitmap;

        internal readonly bool IsEmpty => _storage is null;

        internal readonly bool IsNonEmpty => _storage is not null;

        internal readonly int MaxUsedIndex
        {
            get
            {
                if (_storage is Lut<TElement> lut)
                    return lut.MaxUsedIndex;

                return _bitmap == 0 ? -1 : _bitmap.GetHighestSetBit();
            }
        }

        internal readonly ref readonly TElement Get(int columnIndex)
        {
            if (_storage is TElement[] nodes)
            {
                if (columnIndex >= nodes.Length || (_bitmap & (1u << columnIndex)) == 0)
                    return ref ElementDefault;

                return ref nodes[columnIndex];
            }

            if (_storage is Lut<TElement> lut)
                return ref lut.Get(columnIndex);

            return ref ElementDefault;
        }

        internal readonly bool IsUsed(int columnIndex)
        {
            if (_storage is Lut<TElement> lut)
                return lut.IsUsed(columnIndex);

            return columnIndex < 32 && (_bitmap & (1u << columnIndex)) != 0;
        }

        internal void Set(int columnIndex, TElement value)
        {
            if (_storage is Lut<TElement> lut)
            {
                lut.Set(columnIndex, value);
                if (lut.IsEmpty)
                    _storage = null;
                return;
            }

            if (columnIndex >= 32)
            {
                UpgradeToWideAndSet(columnIndex, value);
                return;
            }

            // Narrow mode
            if (_storage is not TElement[] nodes)
            {
                var size = 4;
                while (columnIndex >= size)
                    size *= 2;

                nodes = new TElement[size];
                _storage = nodes;
            }
            else if (columnIndex >= nodes.Length)
            {
                var size = nodes.Length;
                while (columnIndex >= size)
                    size *= 2;

                Array.Resize(ref nodes, size);
                _storage = nodes;
            }

            nodes[columnIndex] = value;

            var valueIsDefault = EqualityComparer<TElement>.Default.Equals(value, ElementDefault);
            if (valueIsDefault)
                _bitmap &= ~(1u << columnIndex);
            else
                _bitmap |= 1u << columnIndex;

            if (_bitmap == 0)
                _storage = null;
        }

        /// <summary>
        /// Fast path for setting a value that the caller guarantees is not <c>default</c>.
        /// Skips the <see cref="EqualityComparer{TElement}"/> check and always sets the bitmap bit.
        /// </summary>
        internal void SetNonDefault(int columnIndex, TElement value)
        {
            if (_storage is Lut<TElement> lut)
            {
                lut.SetNonDefault(columnIndex, value);
                return;
            }

            if (columnIndex >= 32)
            {
                UpgradeToWideAndSet(columnIndex, value);
                return;
            }

            // Narrow mode
            if (_storage is not TElement[] nodes)
            {
                var size = 4;
                while (columnIndex >= size)
                    size *= 2;

                nodes = new TElement[size];
                _storage = nodes;
            }
            else if (columnIndex >= nodes.Length)
            {
                var size = nodes.Length;
                while (columnIndex >= size)
                    size *= 2;

                Array.Resize(ref nodes, size);
                _storage = nodes;
            }

            nodes[columnIndex] = value;
            _bitmap |= 1u << columnIndex;
        }

        private void UpgradeToWideAndSet(int columnIndex, TElement value)
        {
            var lut = new Lut<TElement>();
            if (_storage is TElement[] existingNodes)
            {
                var bm = _bitmap;
                while (bm != 0)
                {
                    var bit = BitOperations.TrailingZeroCount(bm);
                    lut.Set(bit, existingNodes[bit]);
                    bm &= bm - 1;
                }
            }

            lut.Set(columnIndex, value);
            _storage = lut;
            _bitmap = 0;
        }

        internal static RowData CreateForSet(int columnIndex, TElement value)
        {
            if (columnIndex >= 32)
            {
                var lut = new Lut<TElement>();
                lut.Set(columnIndex, value);
                return new RowData { _storage = lut };
            }

            var size = 4;
            while (columnIndex >= size)
                size *= 2;

            var nodes = new TElement[size];
            nodes[columnIndex] = value;
            return new RowData { _storage = nodes, _bitmap = 1u << columnIndex };
        }

        internal readonly ColumnEnumerator GetColumnEnumerator(int startCol, int endCol)
        {
            if (_storage is TElement[] nodes)
                return new ColumnEnumerator(nodes, _bitmap, startCol, endCol);

            if (_storage is Lut<TElement> lut)
                return new ColumnEnumerator(lut, startCol, endCol);

            return default;
        }

        internal readonly ReverseColumnEnumerator GetReverseColumnEnumerator(int startCol, int endCol)
        {
            if (_storage is TElement[] nodes)
                return new ReverseColumnEnumerator(nodes, _bitmap, startCol, endCol);

            if (_storage is Lut<TElement> lut)
                return new ReverseColumnEnumerator(lut, startCol, endCol);

            return default;
        }

        public readonly bool Equals(RowData other)
            => ReferenceEquals(_storage, other._storage) && _bitmap == other._bitmap;

        public override readonly bool Equals(object? obj)
            => obj is RowData other && Equals(other);

        public override readonly int GetHashCode()
            => HashCode.Combine(_storage, _bitmap);
    }

    /// <summary>
    /// Forward column enumerator that works with both narrow (flat array) and wide (Lut) row storage.
    /// </summary>
    private struct ColumnEnumerator
    {
        private readonly TElement[]? _nodes;
        private uint _remainingBits;
        private Lut<TElement>.LutEnumerator _lutEnumerator;
        private int _idx;
        private readonly bool _isWide;

        internal ColumnEnumerator(TElement[] nodes, uint bitmap, int startCol, int endCol)
        {
            _nodes = nodes;
            _isWide = false;
            _idx = -1;
            _lutEnumerator = default;

            // Mask bitmap to [startCol, endCol]
            var mask = endCol < 31 ? (1u << (endCol + 1)) - 1 : ~0u;
            if (startCol > 0)
                mask &= ~((1u << startCol) - 1);

            _remainingBits = bitmap & mask;
        }

        internal ColumnEnumerator(Lut<TElement> lut, int startCol, int endCol)
        {
            _isWide = true;
            _nodes = null;
            _remainingBits = 0;
            _idx = -1;
            _lutEnumerator = new Lut<TElement>.LutEnumerator(lut, startCol, endCol);
        }

        public ref readonly TElement Current
        {
            get
            {
                if (_isWide)
                    return ref _lutEnumerator.Current;

                return ref _nodes![_idx];
            }
        }

        public int Index
        {
            get
            {
                if (_isWide)
                    return _lutEnumerator.Index;

                return _idx;
            }
        }

        public bool MoveNext()
        {
            if (_isWide)
                return _lutEnumerator.MoveNext();

            if (_remainingBits == 0)
                return false;

            _idx = BitOperations.TrailingZeroCount(_remainingBits);
            _remainingBits &= _remainingBits - 1;
            return true;
        }
    }

    /// <summary>
    /// Reverse column enumerator that works with both narrow (flat array) and wide (Lut) row storage.
    /// </summary>
    private struct ReverseColumnEnumerator
    {
        private readonly TElement[]? _nodes;
        private uint _remainingBits;
        private Lut<TElement>.ReverseLutEnumerator _lutEnumerator;
        private int _idx;
        private readonly bool _isWide;

        internal ReverseColumnEnumerator(TElement[] nodes, uint bitmap, int startCol, int endCol)
        {
            _nodes = nodes;
            _isWide = false;
            _idx = -1;
            _lutEnumerator = default;

            var mask = endCol < 31 ? (1u << (endCol + 1)) - 1 : ~0u;
            if (startCol > 0)
                mask &= ~((1u << startCol) - 1);

            _remainingBits = bitmap & mask;
        }

        internal ReverseColumnEnumerator(Lut<TElement> lut, int startCol, int endCol)
        {
            _isWide = true;
            _nodes = null;
            _remainingBits = 0;
            _idx = -1;
            _lutEnumerator = new Lut<TElement>.ReverseLutEnumerator(lut, startCol, endCol);
        }

        public ref TElement Current
        {
            get
            {
                if (_isWide)
                    return ref _lutEnumerator.Current;

                return ref _nodes![_idx];
            }
        }

        public int Index
        {
            get
            {
                if (_isWide)
                    return _lutEnumerator.Index;

                return _idx;
            }
        }

        public bool MoveNext()
        {
            if (_isWide)
                return _lutEnumerator.MoveNext();

            if (_remainingBits == 0)
                return false;

            _idx = 31 - BitOperations.LeadingZeroCount(_remainingBits);
            _remainingBits &= ~(1u << _idx);
            return true;
        }
    }
}
