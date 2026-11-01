using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel.Coordinates;
using XLibur.Excel.IO;
using XLibur.Extensions;

namespace XLibur.Excel;

public partial class XLWorkbook
{
    /// <summary>
    /// The rich values a save writes, in the order of their vm (value metadata) indexes: one for each
    /// cell with an in-cell image, and then one for each error newer than Excel 2007 that a cell holds,
    /// shared by every cell with that error, as Excel shares it (#847).
    /// </summary>
    /// <remarks>
    /// <para>
    /// A file holds such an error as <c>#VALUE!</c> with a vm index whose rich value names the error
    /// (<see cref="XLErrorExtensions.IsRichValue"/>). An error's index is given when a cell with it is
    /// written, as the result of a formula is known only once the writer has calculated it, and a cell
    /// whose index is not the one it had is given it once the sheets are written.
    /// </para>
    /// <para>
    /// A save copies a sheet it does not change as the file, or the last save, had it (#702, #719),
    /// with the vm index each of its cells had then. So before any sheet is written, every cell that
    /// has a rich value's index takes the one this save gives its error, which changes the sheet when
    /// the index changes, and the sheet is written again. An index of another kind of value metadata,
    /// such as a cube function's <c>XLMDX</c>, is left as it is. A workbook whose rich data XLibur does
    /// not know keeps it as loaded, and its cells their indexes, and has no plan.
    /// </para>
    /// </remarks>
    internal sealed class RichValuePlan
    {
        private readonly WorkbookPart _workbookPart;
        private readonly SaveContext _context;
        private readonly Dictionary<XLError, uint> _errorIndexes = [];
        private readonly List<(XLCellsCollection Cells, Point Point, uint? Index)> _corrections = [];
        private readonly HashSet<uint> _richValueIndexes;
        private uint _richValueTypeIndex;
        private ValueMetadata? _valueMetadata;

        internal RichValuePlan(WorkbookPart workbookPart, SaveContext context)
        {
            _workbookPart = workbookPart;
            _context = context;
            _richValueIndexes = RichValueIndexesOf(workbookPart.CellMetadataPart);
        }

        /// <summary>The images, in the order of their rich values.</summary>
        internal List<RichDataWriter.RichValueEntry> Images { get; } = [];

        /// <summary>The errors, in the order of their rich values, which follow the images'.</summary>
        internal List<XLError> Errors { get; } = [];

        /// <summary>The metadata part's root, once there is a rich value to write.</summary>
        internal Metadata? Metadata { get; private set; }

        /// <summary>Adds the rich value of an image, and returns its vm index.</summary>
        internal uint AddImage(RichDataWriter.RichValueEntry image)
        {
            var index = AppendValueMetadata(Images.Count);
            Images.Add(image);
            return index;
        }

        /// <summary>
        /// Numbers every cell that has a vm index and no image, on every sheet: the index of its error,
        /// or none. Called once the images are numbered and before any sheet is written.
        /// </summary>
        internal void NumberCellsWithIndexes(IEnumerable<XLWorksheet> worksheets)
        {
            var cells = new List<(XLCellsCollection Cells, Point Point, XLError? Error)>();
            foreach (var worksheet in worksheets)
            {
                var cellsCollection = worksheet.Internals.CellsCollection;
                foreach (var point in cellsCollection.GetValueMetaPoints())
                {
                    ref readonly var misc = ref cellsCollection.MiscSlice[point];
                    if (misc.CellImage is not null || !_richValueIndexes.Contains(misc.ValueMetaIndex!.Value))
                        continue;

                    var value = cellsCollection.ValueSlice.GetCellValue(point);
                    cells.Add((cellsCollection, point,
                        value.IsError && value.GetError().IsRichValue() ? value.GetError() : null));
                }
            }

            // In a fixed order, so a save that has the same errors as the last one gives them the same
            // indexes, and a sheet it does not change can stay as that save wrote it.
            foreach (var error in cells.Where(c => c.Error is not null).Select(c => c.Error!.Value).Distinct().Order())
                ErrorIndex(error);

            foreach (var (cellsCollection, point, error) in cells)
                cellsCollection.GetCell(point).ValueMetaIndex = error is { } e ? _errorIndexes[e] : null;
        }

        /// <summary>
        /// The vm index the writer gives a cell that holds <paramref name="value"/>: its image's, its
        /// error's, or none. A cell whose index is not the one it has takes it once the sheets are
        /// written (<see cref="Write"/>).
        /// </summary>
        internal uint? IndexFor(in XLMiscSliceContent misc, in XLCellValue value, XLCellsCollection cells, Point point)
        {
            if (misc.CellImage is not null || (misc.ValueMetaIndex is { } other && !_richValueIndexes.Contains(other)))
                return misc.ValueMetaIndex;

            uint? index = value.IsError && value.GetError().IsRichValue() ? ErrorIndex(value.GetError()) : null;
            if (index != misc.ValueMetaIndex)
                _corrections.Add((cells, point, index));

            return index;
        }

        /// <summary>
        /// Gives each cell the vm index it was written with, and writes the rich data parts and the
        /// rich values' metadata. Called once the sheets are written.
        /// </summary>
        internal void Write(XLInCellImageStore imageStore)
        {
            foreach (var (cellsCollection, point, index) in _corrections)
                cellsCollection.GetCell(point).ValueMetaIndex = index;

            if (Metadata is null)
                return;

            RichDataWriter.WriteRichDataParts(_workbookPart, imageStore, Images, Errors, _context.RelIdGenerator);
            AppendRichValueFutureMetadata(Metadata, Images.Count + Errors.Count);
            PutMetadataChildrenInSchemaOrder(Metadata);
        }

        private uint ErrorIndex(XLError error)
        {
            if (_errorIndexes.TryGetValue(error, out var index))
                return index;

            index = AppendValueMetadata(Images.Count + Errors.Count);
            Errors.Add(error);
            _errorIndexes[error] = index;
            return index;
        }

        /// <summary>
        /// Appends the value metadata of the rich value at <paramref name="richValueIndex"/>, and returns
        /// its vm index, which is 1-based. The first one replaces the rich values' records the part has.
        /// </summary>
        private uint AppendValueMetadata(int richValueIndex)
        {
            if (_valueMetadata is null)
            {
                Metadata = EnsureMetadataPart(_workbookPart, _context);
                _richValueTypeIndex = EnsureRichValueMetadataType(Metadata);
                _valueMetadata = PrepareValueMetadata(Metadata, _richValueTypeIndex);
            }

            var block = new MetadataBlock();
            block.Append(new MetadataRecord { TypeIndex = _richValueTypeIndex, Val = (uint)richValueIndex });
            _valueMetadata.Append(block);
            _valueMetadata.Count = (_valueMetadata.Count ?? 0) + 1;
            var index = _valueMetadata.Count.Value;
            _richValueIndexes.Add(index);
            return index;
        }

        /// <summary>
        /// The vm indexes, 1-based, of the value metadata records of rich values in the metadata part,
        /// as the file or the last save wrote them.
        /// </summary>
        /// <remarks>
        /// The part is read as text unless something has loaded its elements already: loading them
        /// makes the save write the part again from them, and a part this save does not change should
        /// stay as it was.
        /// </remarks>
        private static HashSet<uint> RichValueIndexesOf(CellMetadataPart? part)
        {
            var indexes = new HashSet<uint>();
            if (part is null)
                return indexes;

            if (part.IsRootElementLoaded)
                return RichValueIndexesOf(part.Metadata);

            using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
            var metadata = XDocument.Load(stream).Root;
            XNamespace main = OpenXmlConst.Main2006SsNs;
            var typeIndex = (metadata?.Element(main + "metadataTypes")?.Elements(main + "metadataType") ?? [])
                .ToList().FindIndex(t => (string?)t.Attribute("name") == RichValueMetadataName) + 1;
            var blocks = metadata?.Element(main + "valueMetadata")?.Elements(main + "bk") ?? [];
            uint index = 0;
            foreach (var block in blocks)
            {
                index++;
                if ((string?)block.Element(main + "rc")?.Attribute("t") == typeIndex.ToString(CultureInfo.InvariantCulture))
                    indexes.Add(index);
            }

            return indexes;
        }

        private static HashSet<uint> RichValueIndexesOf(Metadata? metadata)
        {
            var indexes = new HashSet<uint>();
            if (metadata?.MetadataTypes is not { } types || metadata.GetFirstChild<ValueMetadata>() is not { } values)
                return indexes;

            var typeIndex = types.Elements<MetadataType>().ToList().FindIndex(t => t.Name?.Value == RichValueMetadataName) + 1;
            uint index = 0;
            foreach (var block in values.Elements<MetadataBlock>())
            {
                index++;
                if (block.GetFirstChild<MetadataRecord>()?.TypeIndex?.Value == typeIndex)
                    indexes.Add(index);
            }

            return indexes;
        }
    }
}
