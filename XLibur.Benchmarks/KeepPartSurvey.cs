using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel;
using XLibur.Excel.Coordinates;
using XLibur.Excel.IO;

namespace XLibur.Benchmarks;

/// <summary>
/// Loads and saves every workbook under a folder, and reports for each worksheet whether the save
/// kept its part as the file had it (#702 tier 1), kept only its cells (tier 2) or wrote it again,
/// and why. It also reloads each saved workbook and counts the cells whose value or formula changed.
/// </summary>
/// <remarks>
/// Run with: dotnet run -c Release --framework net10.0 --project XLibur.Benchmarks -- profile keepsurvey [folder] [--sheets] [--dump dir]
/// <para>
/// The folder defaults to <c>XLibur.Tests/Resource</c>. <c>--sheets</c> adds one tab-separated
/// line per worksheet after the summary. <c>--dump</c> writes the markup of each sheet not kept
/// whole, without its cells, as the file had it (<c>.file.xml</c>) and as the save wrote it
/// (<c>.save.xml</c>), for a diff.
/// </para>
/// <para>
/// <b>How a kept part is told apart.</b> Before loading, the survey spells the XML declaration of
/// every worksheet part as no save writes it, as <c>KeepUnchangedPartTests</c> does. The declaration
/// does not take part in the tier-1 comparison, so the save decides as it would have. A saved part
/// that is byte for byte the marked source was kept whole. A saved part whose cells are byte for
/// byte the source's had its cells kept; for a file XLibur wrote, a rewrite can match too, so that
/// column is an upper bound there.
/// </para>
/// <para>
/// <b>Why a part was not kept.</b> The survey cuts the cells out of the source part and the saved
/// part and compares the rest as tier 1 does (<see cref="XmlInfoset"/>). When they differ, the first
/// difference names the category, such as <c>worksheet@mc:Ignorable added</c>. When they match,
/// the sheet was excluded from the fast path, and the survey names the exclusion it can see in the
/// model: formulas, pivot tables, a totals row, cell metadata. A sheet with none of those was
/// excluded for the whole workbook, by a styles part that was not loaded or shared strings the save
/// does not write at their file indices.
/// </para>
/// <para>
/// The producer of a file is read from its <c>docProps/app.xml</c>: Excel, XLibur or ClosedXML, or
/// other (see <see cref="ProducerOf"/>).
/// </para>
/// </remarks>
public static class KeepPartSurvey
{
    /// <summary>The XML declaration as no save writes it.</summary>
    private const string Marker = "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>";

    private const string Excel = "Excel";
    private const string XLiburOrClosedXml = "XLibur/ClosedXML";
    private const string Other = "Other";

    private static readonly string[] Producers = [Excel, XLiburOrClosedXml, Other];

    /// <summary>Where to write the markup of each sheet not kept whole, or null.</summary>
    private static string? _dumpFolder;

    public static void Run(string[] args)
    {
        var options = args.Skip(2).ToList();
        var showSheets = options.Remove("--sheets");
        var dump = options.IndexOf("--dump");
        if (dump >= 0 && dump + 1 < options.Count)
        {
            _dumpFolder = Directory.CreateDirectory(options[dump + 1]).FullName;
            options.RemoveRange(dump, 2);
        }

        var folder = options.Count > 0 ? options[0] : FindResourceFolder();

        var files = Directory.EnumerateFiles(folder, "*.*", SearchOption.AllDirectories)
            .Where(f => f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)
                        || f.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase))
            .OrderBy(f => f, StringComparer.Ordinal)
            .ToList();

        var sheets = new List<SheetResult>();
        var failures = new List<string>();
        var notLoaded = 0;
        foreach (var file in files)
        {
            var name = Path.GetRelativePath(folder, file).Replace('\\', '/');
            try
            {
                sheets.AddRange(Survey(name, File.ReadAllBytes(file), out var loaded));
                if (!loaded)
                    notLoaded++;
            }
            catch (Exception e)
            {
                failures.Add($"{name}: {e.GetType().Name}: {FirstLine(e.Message)}");
            }
        }

        Report(folder, files.Count, notLoaded, sheets, failures);
        if (showSheets)
            ReportSheets(sheets);
    }

    /// <summary>
    /// Loads, saves and reloads one workbook. Returns no sheets, with <paramref name="loaded"/>
    /// false, when the source does not load; a source that loads but does not save or reload throws.
    /// </summary>
    private static List<SheetResult> Survey(string file, byte[] original, out bool loaded)
    {
        loaded = false;
        byte[] source;
        string producer;
        try
        {
            source = Mark(original, out producer);
        }
        catch (Exception)
        {
            return [];
        }

        XLWorkbook workbook;
        try
        {
            workbook = new XLWorkbook(new MemoryStream(source, writable: false));
        }
        catch (Exception)
        {
            return [];
        }

        loaded = true;
        using (workbook)
        {
            var exclusions = workbook.Worksheets.Cast<XLWorksheet>()
                .ToDictionary(ws => ws.Name, ExclusionsOf, StringComparer.OrdinalIgnoreCase);

            byte[] saved;
            using (var output = new MemoryStream())
            {
                workbook.SaveAs(output);
                saved = output.ToArray();
            }

            using var reloaded = new XLWorkbook(new MemoryStream(saved, writable: false));
            using var sourcePackage = SpreadsheetDocument.Open(new MemoryStream(source, writable: false), false);
            using var savedPackage = SpreadsheetDocument.Open(new MemoryStream(saved, writable: false), false);
            var savedParts = WorksheetParts(savedPackage);

            var results = new List<SheetResult>();
            foreach (var (sheetName, sourcePart) in WorksheetParts(sourcePackage))
            {
                if (!savedParts.TryGetValue(sheetName, out var savedPart)
                    || !workbook.TryGetWorksheet(sheetName, out IXLWorksheet? sheet))
                {
                    continue;
                }

                var mismatches = reloaded.TryGetWorksheet(sheetName, out IXLWorksheet? reloadedSheet)
                    ? CountValueMismatches(sheet, reloadedSheet)
                    : sheet.CellsUsed(XLCellsUsedOptions.Contents).Count();
                var excluded = exclusions.TryGetValue(sheetName, out var e) ? e : string.Empty;
                results.Add(Classify(file, sheetName, producer, sourcePart, savedPart, excluded, mismatches));
            }

            return results;
        }
    }

    private static SheetResult Classify(string file, string sheet, string producer, WorksheetPart sourcePart,
        WorksheetPart savedPart, string exclusions, int mismatches)
    {
        var sourceBytes = ReadAll(sourcePart);
        var savedBytes = ReadAll(savedPart);
        var result = new SheetResult(file, sheet, producer, Outcome.Rewritten, string.Empty, string.Empty,
            exclusions, mismatches);

        if (sourceBytes.AsSpan().SequenceEqual(savedBytes) && StartsWithMarker(savedBytes))
            return result with { Outcome = Outcome.KeptWhole };

        using var sourceBuffer = WorksheetPartBuffer.TryRead(sourcePart);
        using var savedBuffer = WorksheetPartBuffer.TryRead(savedPart);
        if (sourceBuffer is null || savedBuffer is null)
            return result with { Category = "part not buffered (larger than 16 MB, or no sheetData found)" };

        var outcome = CellsMatch(sourceBuffer, savedBuffer) ? Outcome.CellsKept : Outcome.Rewritten;
        var sourceMarkup = sourceBuffer.CopyWithoutSheetData();
        var savedMarkup = savedBuffer.CopyWithoutSheetData();
        if (_dumpFolder is not null)
        {
            var stem = Path.Combine(_dumpFolder,
                string.Concat($"{file}__{sheet}".Select(c => char.IsLetterOrDigit(c) || c is '-' or '.' ? c : '_')));
            File.WriteAllBytes(stem + ".file.xml", sourceMarkup);
            File.WriteAllBytes(stem + ".save.xml", savedMarkup);
        }

        bool equivalent;
        using (var left = new MemoryStream(sourceMarkup, writable: false))
        using (var right = new MemoryStream(savedMarkup, writable: false))
            equivalent = XmlInfoset.AreEquivalent(left, right);

        if (!equivalent)
        {
            var (category, detail) = MarkupDifference.First(sourceMarkup, savedMarkup);
            return result with { Outcome = outcome, Category = "markup: " + category, Detail = detail };
        }

        var reason = exclusions.Length > 0
            ? "excluded: " + exclusions
            : "excluded for the workbook: styles not loaded, or shared strings not at file indices";
        return result with { Outcome = outcome, Category = reason };
    }

    /// <summary>
    /// The exclusions from the fast path that the model shows, as <c>LoadedSheetData</c> lists them.
    /// </summary>
    private static string ExclusionsOf(XLWorksheet sheet)
    {
        var reasons = new List<string>();
        var cells = sheet.Internals.CellsCollection;
        if (!cells.FormulaSlice.IsEmpty)
            reasons.Add("formulas");
        if (sheet.PivotTables.Any<IXLPivotTable>())
            reasons.Add("pivot table");
        if (sheet.Tables.Any<IXLTable>(t => t.ShowTotalsRow))
            reasons.Add("totals row");
        if (HasCellMetadata(cells.MiscSlice))
            reasons.Add("cell metadata");
        if (sheet.CellsChangedSinceLoad())
            reasons.Add("cells changed by the load");
        return string.Join(", ", reasons);
    }

    private static bool HasCellMetadata(Slice<XLMiscSliceContent> misc)
    {
        var enumerator = new Slice<XLMiscSliceContent>.Enumerator(misc, Area.Full);
        while (enumerator.MoveNext())
        {
            ref readonly var content = ref enumerator.Current;
            if (content.CellMetaIndex is not null || content.ValueMetaIndex is not null || content.CellImage is not null)
                return true;
        }

        return false;
    }

    /// <summary>
    /// The cells whose cached value or formula differs between the saved model and its reload, or
    /// that only one of them has.
    /// </summary>
    private static int CountValueMismatches(IXLWorksheet saved, IXLWorksheet reloaded)
    {
        var reloadedCells = new Dictionary<string, (XLCellValue Value, string Formula)>();
        foreach (var cell in reloaded.CellsUsed(XLCellsUsedOptions.Contents))
            reloadedCells[cell.Address.ToString()] = (cell.CachedValue, Formula(cell));

        var mismatches = 0;
        foreach (var cell in saved.CellsUsed(XLCellsUsedOptions.Contents))
        {
            if (!reloadedCells.Remove(cell.Address.ToString(), out var other)
                || !other.Value.Equals(cell.CachedValue)
                || other.Formula != Formula(cell))
            {
                mismatches++;
            }
        }

        return mismatches + reloadedCells.Count;
    }

    private static string Formula(IXLCell cell) => cell.HasFormula ? cell.FormulaA1 : string.Empty;

    /// <summary>
    /// Is the content of <c>&lt;sheetData&gt;</c>, after its start tag, the same bytes in both? The
    /// start tag can gain namespace declarations when the cells are kept.
    /// </summary>
    private static bool CellsMatch(WorksheetPartBuffer source, WorksheetPartBuffer saved)
    {
        if (source.SheetData.IsEmptyElement || saved.SheetData.IsEmptyElement)
            return false;

        var left = source.Xml[source.SheetData.StartTagEnd..source.SheetData.End];
        var right = saved.Xml[saved.SheetData.StartTagEnd..saved.SheetData.End];
        return left.SequenceEqual(right);
    }

    /// <summary>
    /// A copy of the package with the XML declaration of every worksheet part spelled as
    /// <see cref="Marker"/>, and the producer of the file.
    /// </summary>
    private static byte[] Mark(byte[] original, out string producer)
    {
        List<string> partNames;
        using (var package = SpreadsheetDocument.Open(new MemoryStream(original, writable: false), false))
        {
            producer = ProducerOf(package);
            partNames = WorksheetParts(package).Values
                .Select(part => part.Uri.OriginalString.TrimStart('/'))
                .ToList();
        }

        using var copy = new MemoryStream();
        copy.Write(original);
        using (var zip = new ZipArchive(copy, ZipArchiveMode.Update, leaveOpen: true))
        {
            foreach (var partName in partNames)
            {
                var entry = zip.Entries.FirstOrDefault(e =>
                    string.Equals(e.FullName, partName, StringComparison.OrdinalIgnoreCase));
                if (entry is null)
                    continue;

                byte[] xml;
                using (var stream = entry.Open())
                using (var buffer = new MemoryStream())
                {
                    stream.CopyTo(buffer);
                    xml = buffer.ToArray();
                }

                var marked = WithMarker(xml);
                if (marked is null)
                    continue;

                using var output = entry.Open();
                output.SetLength(0);
                output.Write(marked);
            }
        }

        return copy.ToArray();
    }

    /// <summary>
    /// The part with its declaration replaced by <see cref="Marker"/>, or null when the part is not
    /// UTF-8.
    /// </summary>
    private static byte[]? WithMarker(byte[] xml)
    {
        var span = xml.AsSpan();
        if (span is [0xEF, 0xBB, 0xBF, ..])
            span = span[3..];
        else if (span.Length >= 2 && (span[0] == 0xFF || span[0] == 0xFE || span[0] == 0))
            return null;

        if (span.StartsWith("<?xml"u8))
        {
            var end = span.IndexOf("?>"u8);
            if (end < 0)
                return null;

            var declaration = Encoding.ASCII.GetString(span[..end]);
            if (declaration.Contains("encoding", StringComparison.Ordinal)
                && !declaration.Contains("utf-8", StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            span = span[(end + 2)..];
        }

        var marker = Encoding.ASCII.GetBytes(Marker);
        var result = new byte[marker.Length + span.Length];
        marker.CopyTo(result, 0);
        span.CopyTo(result.AsSpan(marker.Length));
        return result;
    }

    private static bool StartsWithMarker(byte[] xml) => xml.AsSpan().StartsWith(Encoding.ASCII.GetBytes(Marker));

    /// <summary>
    /// XLibur and ClosedXML write <c>Microsoft Excel</c> as the application too, but through the
    /// OpenXML SDK, which spells the root of the part <c>&lt;ap:Properties&gt;</c>. Excel declares
    /// the namespace as the default.
    /// </summary>
    private static string ProducerOf(SpreadsheetDocument package)
    {
        var part = package.ExtendedFilePropertiesPart;
        if (part is null)
            return Other;

        string prefix;
        using (var stream = part.GetStream(FileMode.Open, FileAccess.Read))
        using (var reader = PartXmlReader.Create(stream))
            prefix = reader.MoveToContent() == XmlNodeType.Element ? reader.Prefix : string.Empty;

        var application = part.Properties?.Application?.Text ?? string.Empty;
        if (prefix == "ap"
            || application.Contains("XLibur", StringComparison.OrdinalIgnoreCase)
            || application.Contains("ClosedXML", StringComparison.OrdinalIgnoreCase))
            return XLiburOrClosedXml;

        return application.Contains("Excel", StringComparison.OrdinalIgnoreCase) ? Excel : Other;
    }

    /// <summary>The worksheet parts of the package by sheet name; chartsheets and others left out.</summary>
    private static Dictionary<string, WorksheetPart> WorksheetParts(SpreadsheetDocument package)
    {
        var parts = new Dictionary<string, WorksheetPart>(StringComparer.OrdinalIgnoreCase);
        var workbookPart = package.WorkbookPart;
        var sheets = workbookPart?.Workbook?.Sheets;
        if (workbookPart is null || sheets is null)
            return parts;

        foreach (var sheet in sheets.Elements<Sheet>())
        {
            var id = sheet.Id?.Value;
            var name = sheet.Name?.Value;
            if (id is null || name is null)
                continue;

            if (workbookPart.TryGetPartById(id, out var part) && part is WorksheetPart worksheetPart)
                parts.TryAdd(name, worksheetPart);
        }

        return parts;
    }

    private static byte[] ReadAll(OpenXmlPart part)
    {
        using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return buffer.ToArray();
    }

    private static void Report(string folder, int fileCount, int notLoaded, List<SheetResult> sheets,
        List<string> failures)
    {
        Console.WriteLine($"Keep-part survey of {folder}");
        Console.WriteLine($"{fileCount} workbooks: {notLoaded} do not load, {failures.Count} fail to save or reload.");
        Console.WriteLine();
        Console.WriteLine("Producer           Sheets  Kept whole  Cells kept  Rewritten  Mismatched cells  Sheets with mismatches");
        foreach (var producer in Producers)
        {
            var group = sheets.Where(s => s.Producer == producer).ToList();
            Console.WriteLine(string.Create(CultureInfo.InvariantCulture,
                $"{producer,-18} {group.Count,6}  {Count(group, Outcome.KeptWhole),10}  {Count(group, Outcome.CellsKept),10}  "
                + $"{Count(group, Outcome.Rewritten),9}  {group.Sum(s => s.ValueMismatches),16}  {group.Count(s => s.ValueMismatches > 0),22}"));
        }

        foreach (var producer in Producers)
            ReportMisses(producer, sheets.Where(s => s.Producer == producer && s.Outcome != Outcome.KeptWhole).ToList());

        ReportExclusions(sheets.Where(s => s.Exclusions.Length > 0).ToList());

        if (failures.Count > 0)
        {
            Console.WriteLine();
            Console.WriteLine("Failed to save or reload:");
            foreach (var failure in failures)
                Console.WriteLine("  " + failure);
        }
    }

    /// <summary>The sheets of one producer not kept whole, by category, with an example of each.</summary>
    private static void ReportMisses(string producer, List<SheetResult> misses)
    {
        if (misses.Count == 0)
            return;

        Console.WriteLine();
        Console.WriteLine($"Not kept whole, {producer}: {misses.Count}");
        foreach (var category in misses.GroupBy(s => s.Category)
                     .OrderByDescending(g => g.Count()).ThenBy(g => g.Key, StringComparer.Ordinal))
        {
            var example = category.First();
            var detail = example.Detail.Length > 0 ? $" ({Truncate(example.Detail, 120)})" : string.Empty;
            Console.WriteLine($"{category.Count(),6}  {category.Key}");
            Console.WriteLine($"        e.g. {example.File} [{example.Sheet}]{detail}");
        }
    }

    /// <summary>
    /// For the issues that lift an exclusion: how many excluded sheets would then be kept whole,
    /// because the rest of their markup already matches?
    /// </summary>
    private static void ReportExclusions(List<SheetResult> excluded)
    {
        if (excluded.Count == 0)
            return;

        Console.WriteLine();
        Console.WriteLine("Excluded sheets by exclusion: markup equal / markup differs");
        var reasons = excluded.SelectMany(s => s.Exclusions.Split(", ")).Distinct().Order(StringComparer.Ordinal);
        foreach (var (reason, producer) in reasons.SelectMany(reason => Producers.Select(producer => (reason, producer))))
        {
            var group = excluded.Where(s => s.Producer == producer && s.Exclusions.Split(", ").Contains(reason)).ToList();
            if (group.Count == 0)
                continue;

            var equal = group.Count(s => !s.Category.StartsWith("markup: ", StringComparison.Ordinal));
            Console.WriteLine($"  {reason,-26} {producer,-18} {equal,5} / {group.Count - equal}");
        }
    }

    private static void ReportSheets(List<SheetResult> sheets)
    {
        Console.WriteLine();
        Console.WriteLine("# file\tsheet\tproducer\toutcome\tmismatches\texclusions\tcategory\tdetail");
        foreach (var s in sheets)
        {
            Console.WriteLine(string.Join('\t', s.File, s.Sheet, s.Producer, s.Outcome,
                s.ValueMismatches.ToString(CultureInfo.InvariantCulture), s.Exclusions, s.Category, s.Detail));
        }
    }

    private static int Count(List<SheetResult> sheets, Outcome outcome) => sheets.Count(s => s.Outcome == outcome);

    private static string Truncate(string text, int length) => text.Length <= length ? text : text[..length] + "…";

    private static string FirstLine(string text)
    {
        var newline = text.IndexOfAny(['\r', '\n']);
        return newline < 0 ? text : text[..newline];
    }

    private static string FindResourceFolder()
    {
        for (var dir = new DirectoryInfo(AppContext.BaseDirectory); dir is not null; dir = dir.Parent)
        {
            var candidate = Path.Combine(dir.FullName, "XLibur.Tests", "Resource");
            if (Directory.Exists(candidate))
                return candidate;
        }

        throw new DirectoryNotFoundException("XLibur.Tests/Resource not found above " + AppContext.BaseDirectory);
    }

    private enum Outcome
    {
        KeptWhole,
        CellsKept,
        Rewritten,
    }

    private sealed record SheetResult(
        string File,
        string Sheet,
        string Producer,
        Outcome Outcome,
        string Category,
        string Detail,
        string Exclusions,
        int ValueMismatches);

    /// <summary>
    /// Finds the first place where two documents differ by the rules of <see cref="XmlInfoset"/>, and
    /// names it by the element path and what differs, leaving values out of the name so sheets group.
    /// </summary>
    private static class MarkupDifference
    {
        private enum Kind
        {
            Start,
            End,
            Text,
        }

        private sealed record Token(Kind Kind, string Path, string Name, Dictionary<string, Attribute> Attributes,
            string Text);

        private sealed record Attribute(string Name, string Value, string Resolved);

        /// <summary>The category and a detail with the values, the file on the left.</summary>
        internal static (string Category, string Detail) First(byte[] file, byte[] save)
        {
            var left = Tokens(file);
            var right = Tokens(save);
            var common = Math.Min(left.Count, right.Count);
            for (var i = 0; i < common; i++)
            {
                if (TokenDifference(left[i], right[i]) is { } difference)
                    return difference;
            }

            if (left.Count == right.Count)
                return ("none found", string.Empty);

            return left.Count > common
                ? ($"{left[common].Path} dropped", string.Empty)
                : ($"{right[common].Path} added", string.Empty);
        }

        /// <summary>How <paramref name="save"/> differs from <paramref name="file"/>, or null.</summary>
        private static (string, string)? TokenDifference(Token file, Token save)
        {
            if (file.Kind != save.Kind)
                return KindDifference(file, save);

            return file.Kind switch
            {
                Kind.Start when file.Name != save.Name => (
                    $"{Parent(file.Path)}: save writes <{LocalName(save.Path)}> where the file has <{LocalName(file.Path)}>",
                    string.Empty),
                Kind.Start => AttributeDifference(file, save),
                Kind.Text when file.Text != save.Text => ($"{file.Path} text", $"file={file.Text}, save={save.Text}"),
                _ => null,
            };
        }

        /// <summary>One document has an element or text where the other has something else.</summary>
        private static (string, string) KindDifference(Token file, Token save)
        {
            if (save.Kind == Kind.Start)
                return ($"{save.Path} added", string.Empty);
            if (file.Kind == Kind.Start)
                return ($"{file.Path} dropped", string.Empty);
            return ($"{(file.Kind == Kind.Text ? file.Path : save.Path)} text", string.Empty);
        }

        private static (string, string)? AttributeDifference(Token file, Token save)
        {
            foreach (var (key, attribute) in save.Attributes)
            {
                if (!file.Attributes.ContainsKey(key))
                    return ($"{save.Path}@{attribute.Name} added", $"{attribute.Name}={attribute.Value}");
            }

            foreach (var (key, attribute) in file.Attributes)
            {
                if (!save.Attributes.TryGetValue(key, out var saved))
                    return ($"{file.Path}@{attribute.Name} dropped", $"{attribute.Name}={attribute.Value}");
                if (saved.Value != attribute.Value)
                    return ($"{file.Path}@{attribute.Name} changed", $"file={attribute.Value}, save={saved.Value}");
                if (saved.Resolved != attribute.Resolved)
                    return ($"{file.Path}@{attribute.Name} prefixes resolve otherwise",
                        $"file={attribute.Resolved}, save={saved.Resolved}");
            }

            return null;
        }

        private static List<Token> Tokens(byte[] xml)
        {
            var tokens = new List<Token>();
            var path = new List<string>();
            using var stream = new MemoryStream(xml, writable: false);
            using var reader = PartXmlReader.Create(stream);
            var text = new StringBuilder();

            void FlushText()
            {
                if (text.Length == 0)
                    return;
                tokens.Add(new Token(Kind.Text, string.Join('/', path), string.Empty, [], text.ToString()));
                text.Clear();
            }

            while (reader.Read())
            {
                switch (reader.NodeType)
                {
                    case XmlNodeType.Element:
                        FlushText();
                        path.Add(reader.LocalName);
                        var name = "{" + reader.NamespaceURI + "}" + reader.LocalName;
                        var empty = reader.IsEmptyElement;
                        tokens.Add(new Token(Kind.Start, string.Join('/', path), name, ReadAttributes(reader),
                            string.Empty));
                        if (empty)
                        {
                            tokens.Add(new Token(Kind.End, string.Join('/', path), name, [], string.Empty));
                            path.RemoveAt(path.Count - 1);
                        }

                        break;
                    case XmlNodeType.EndElement:
                        FlushText();
                        tokens.Add(new Token(Kind.End, string.Join('/', path), string.Empty, [], string.Empty));
                        path.RemoveAt(path.Count - 1);
                        break;
                    case XmlNodeType.Text:
                    case XmlNodeType.CDATA:
                    case XmlNodeType.SignificantWhitespace:
                        text.Append(reader.Value);
                        break;
                }
            }

            return tokens;
        }

        private static Dictionary<string, Attribute> ReadAttributes(XmlReader reader)
        {
            var attributes = new Dictionary<string, Attribute>(StringComparer.Ordinal);
            if (!reader.MoveToFirstAttribute())
                return attributes;

            do
            {
                if (reader.NamespaceURI == "http://www.w3.org/2000/xmlns/")
                    continue;

                var name = reader.Prefix.Length > 0 ? reader.Prefix + ":" + reader.LocalName : reader.LocalName;
                var resolved = ListsPrefixes(reader) ? Resolve(reader, reader.Value) : string.Empty;
                attributes["{" + reader.NamespaceURI + "}" + reader.LocalName] = new Attribute(name, reader.Value, resolved);
            } while (reader.MoveToNextAttribute());

            reader.MoveToElement();
            return attributes;
        }

        private static bool ListsPrefixes(XmlReader attribute) =>
            attribute.NamespaceURI == "http://schemas.openxmlformats.org/markup-compatibility/2006"
            || (attribute.NamespaceURI.Length == 0 && attribute.LocalName == "Requires");

        private static string Resolve(XmlReader reader, string value) =>
            string.Join(' ', value.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries).Select(token =>
            {
                var colon = token.IndexOf(':');
                var prefix = colon < 0 ? token : token[..colon];
                return prefix + "=" + reader.LookupNamespace(prefix);
            }));

        private static string Parent(string path)
        {
            var slash = path.LastIndexOf('/');
            return slash < 0 ? string.Empty : path[..slash];
        }

        private static string LocalName(string path) => path[(path.LastIndexOf('/') + 1)..];
    }
}
