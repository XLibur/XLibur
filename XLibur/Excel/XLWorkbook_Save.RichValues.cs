using System;
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
    /// such as a cube function's <c>XLMDX</c>, is left as it is, and a rich value's record is taken
    /// where it is, so no other record moves.
    /// </para>
    /// <para>
    /// A workbook whose rich data XLibur does not know keeps it as loaded, and each cell its index. Its
    /// plan adds the rich value of each newer error a cell without an index holds after the ones the
    /// file has, which keep their indexes.
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

        /// <summary>Whether the workbook keeps the rich data it loaded (<see cref="KeepsLoadedRichData"/>).</summary>
        private readonly bool _keepsLoadedRichData;

        /// <summary>The kept rich values part, which new errors are added to.</summary>
        private readonly OpenXmlPart? _keptRichValuePart;

        /// <summary>The kept structures part, which the structure of an error is added to.</summary>
        private readonly OpenXmlPart? _keptStructurePart;

        /// <summary>How many rich values the kept rich data has.</summary>
        private int _keptRichValueCount;

        /// <summary>How many future metadata blocks of rich values the kept rich data has.</summary>
        private int _keptFutureBlockCount;

        /// <summary>
        /// The records of rich values the part has, by vm index, not yet given to one of this save's.
        /// </summary>
        private readonly Queue<(uint Index, MetadataRecord Record)> _freeRecords = [];

        internal RichValuePlan(WorkbookPart workbookPart, SaveContext context, bool keepsLoadedRichData)
        {
            _workbookPart = workbookPart;
            _context = context;
            _keepsLoadedRichData = keepsLoadedRichData;
            _richValueIndexes = RichValueIndexesOf(workbookPart.CellMetadataPart);
            if (keepsLoadedRichData)
            {
                _keptRichValuePart = FindPart(workbookPart, RichDataWriter.RichValueRelType);
                _keptStructurePart = FindPart(workbookPart, RichDataWriter.RichValueStructureRelType);
            }
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
            if (misc.CellImage is not null || (misc.ValueMetaIndex is { } other && (_keepsLoadedRichData || !_richValueIndexes.Contains(other))))
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

            if (_keepsLoadedRichData)
            {
                AddErrorsToKeptRichData(Metadata);
                return;
            }

            // An old record no rich value of this save took names the first rich value, as no cell
            // names it any more, and removing it would move the records after it.
            while (_freeRecords.TryDequeue(out var free))
                free.Record.Val = 0;

            RichDataWriter.WriteRichDataParts(_workbookPart, imageStore, Images, Errors, _context.RelIdGenerator);
            AppendRichValueFutureMetadata(Metadata, Images.Count + Errors.Count);
            PutMetadataChildrenInSchemaOrder(Metadata);
        }

        /// <summary>
        /// The vm index of <paramref name="error"/>'s rich value, given at its first cell. Null when the
        /// workbook keeps rich data that has no rich values or structures part to add it to.
        /// </summary>
        private uint? ErrorIndex(XLError error)
        {
            if (_errorIndexes.TryGetValue(error, out var index))
                return index;

            if (_keepsLoadedRichData && (_keptRichValuePart is null || _keptStructurePart is null))
                return null;

            // A value metadata record names a future metadata block, which names the rich value. A save
            // that writes the rich data numbers them alike; kept rich data has blocks of its own first.
            if (_valueMetadata is null)
                PrepareValueMetadata();

            index = AppendValueMetadata((_keepsLoadedRichData ? _keptFutureBlockCount : Images.Count) + Errors.Count);
            Errors.Add(error);
            _errorIndexes[error] = index;
            return index;
        }

        /// <summary>
        /// Adds the rich value of each error after the kept ones: an <c>_error</c> structure after the
        /// structures, the rich values after theirs, and their future metadata blocks after theirs.
        /// </summary>
        private void AddErrorsToKeptRichData(Metadata metadata)
        {
            var structureIndex = AppendElements(_keptStructurePart!, "s", ns =>
                [new XElement(ns + "s", new XAttribute("t", RichDataWriter.ErrorStructureType),
                    new XElement(ns + "k", new XAttribute("n", RichDataWriter.ErrorTypeKey), new XAttribute("t", "i")))]);
            AppendElements(_keptRichValuePart!, "rv", ns => Errors.Select(error =>
                new XElement(ns + "rv", new XAttribute("s", structureIndex), new XElement(ns + "v", (int)error))));

            var futureMetadata = metadata.Elements<FutureMetadata>().FirstOrDefault(fm => fm.Name?.Value == RichValueMetadataName);
            if (futureMetadata is null)
            {
                futureMetadata = new FutureMetadata { Name = RichValueMetadataName, Count = 0 };
                metadata.Append(futureMetadata);
            }

            for (var i = 0; i < Errors.Count; i++)
                futureMetadata.Append(NewRichValueFutureBlock(_keptRichValueCount + i));

            futureMetadata.Count = (uint)futureMetadata.Elements<FutureMetadataBlock>().Count();
            PutMetadataChildrenInSchemaOrder(metadata);
        }

        /// <summary>
        /// Appends the elements <paramref name="create"/> makes, in the part's namespace, to the root of
        /// <paramref name="part"/>, and sets its <c>count</c>. Returns how many elements named
        /// <paramref name="name"/> the root had before.
        /// </summary>
        private static int AppendElements(OpenXmlPart part, string name, Func<XNamespace, IEnumerable<XElement>> create)
        {
            XDocument document;
            using (var stream = part.GetStream(FileMode.Open, FileAccess.Read))
                document = XDocument.Load(stream);

            var root = document.Root!;
            var ns = root.Name.Namespace;
            var before = root.Elements(ns + name).Count();
            root.Add(create(ns));
            root.SetAttributeValue("count", root.Elements(ns + name).Count());

            using var buffer = new MemoryStream();
            document.Save(buffer, System.Xml.Linq.SaveOptions.DisableFormatting);
            buffer.Position = 0;
            part.FeedData(buffer);
            return before;
        }

        private static int CountElements(OpenXmlPart part, string name)
        {
            using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
            var root = XDocument.Load(stream).Root!;
            return root.Elements(root.Name.Namespace + name).Count();
        }

        private static OpenXmlPart? FindPart(WorkbookPart workbookPart, string relationshipType)
        {
            foreach (var idPart in workbookPart.Parts)
            {
                if (idPart.OpenXmlPart.RelationshipType == relationshipType)
                    return idPart.OpenXmlPart;
            }

            return null;
        }

        /// <summary>
        /// Gives the rich value at <paramref name="richValueIndex"/> a value metadata record, and returns
        /// its vm index, which is 1-based. A record of a rich value the part has is taken first, where it
        /// is, so a record of another kind, such as a cube function's <c>XLMDX</c>, keeps its index; a
        /// new one is appended.
        /// </summary>
        private uint AppendValueMetadata(int richValueIndex)
        {
            if (_valueMetadata is null)
                PrepareValueMetadata();

            if (_freeRecords.TryDequeue(out var free))
            {
                free.Record.Val = (uint)richValueIndex;
                return free.Index;
            }

            var block = new MetadataBlock();
            block.Append(new MetadataRecord { TypeIndex = _richValueTypeIndex, Val = (uint)richValueIndex });
            _valueMetadata!.Append(block);
            _valueMetadata.Count = (_valueMetadata.Count ?? 0) + 1;
            var index = _valueMetadata.Count.Value;
            _richValueIndexes.Add(index);
            return index;
        }

        private void PrepareValueMetadata()
        {
            Metadata = EnsureMetadataPart(_workbookPart, _context);
            _richValueTypeIndex = EnsureRichValueMetadataType(Metadata);
            _valueMetadata = Metadata.GetFirstChild<ValueMetadata>();
            if (_keepsLoadedRichData)
            {
                // Kept rich values keep their records; new ones are appended after them.
                _keptRichValueCount = CountElements(_keptRichValuePart!, "rv");
                _keptFutureBlockCount = Metadata.Elements<FutureMetadata>()
                    .FirstOrDefault(fm => fm.Name?.Value == RichValueMetadataName)?.Elements<FutureMetadataBlock>().Count() ?? 0;
            }

            if (_valueMetadata is null)
            {
                _valueMetadata = new ValueMetadata { Count = 0 };
                Metadata.Append(_valueMetadata);
                return;
            }

            if (_keepsLoadedRichData)
                return;

            uint index = 0;
            foreach (var block in _valueMetadata.Elements<MetadataBlock>())
            {
                index++;
                if (block.GetFirstChild<MetadataRecord>() is { } record && record.TypeIndex?.Value == _richValueTypeIndex)
                    _freeRecords.Enqueue((index, record));
            }
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
