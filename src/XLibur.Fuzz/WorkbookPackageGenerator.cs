using System.Globalization;
using System.IO.Compression;
using System.Text;

namespace XLibur.Fuzz;

/// <summary>
/// Builds an <c>.xlsx</c> package from fuzzer bytes. Every package it produces is a valid ZIP
/// containing well-formed XML; what varies is whether the workbook <em>described</em> by that XML
/// makes sense.
///
/// <para>
/// This exists because blind mutation of an existing <c>.xlsx</c> barely reaches XLibur at all.
/// An <c>.xlsx</c> is a ZIP with per-entry CRCs, so a flipped byte inside a compressed stream is
/// rejected on checksum by the packaging layer before any XLibur reader runs. After a week of
/// blind fuzzing the corpus held 23 entries of exactly the seed's length and 8 truncations, and
/// not one structurally different package: almost every mutation died at the container and
/// returned no new coverage for the fuzzer to follow.
/// </para>
///
/// <para>
/// So the bytes drive a generator instead of the file. The container is always correct, which
/// spends the fuzzer's budget on the questions that are actually XLibur's: dangling relationship
/// ids, style indices past the end of the table, shared-string indices past the end of the table,
/// a declared dimension that disagrees with the cells present, duplicate sheet names. That band is
/// chosen deliberately — it is where the open defect register already clusters, and it is
/// unreachable from the blind target.
/// </para>
/// </summary>
internal static class WorkbookPackageGenerator
{
    private const string SpreadsheetMlNamespace =
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    private const string RelationshipsNamespace =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    private const string PackageRelationshipsNamespace =
        "http://schemas.openxmlformats.org/package/2006/relationships";

    /// <summary>How many strings the generated shared-string table holds.</summary>
    private const int SharedStringCount = 3;

    /// <summary>How many entries the generated cell-format table holds.</summary>
    private const int CellFormatCount = 2;

    /// <summary>
    /// One in how many chances a given field is generated in a hostile rather than an ordinary
    /// shape.
    ///
    /// <para>
    /// This is not a knob for taste; it is load-bearing, and the first run proved it. With hostile
    /// shapes at even odds, a package with up to three sheets carried a dangling relationship id
    /// about seven times in eight — and a dangling id stops the load immediately (D28), so 30 of
    /// the first 31 generated packages ended at the same line and nothing behind it was reachable
    /// at all. A fuzzer that always trips the first hurdle explores nothing past it.
    /// </para>
    ///
    /// <para>
    /// The point of a structure-aware generator is that most of what it emits must be *loadable*,
    /// so the interesting inputs are the ones carrying a single oddity deep inside an otherwise
    /// ordinary workbook. Raising this number makes packages more ordinary and defects rarer per
    /// input but reachable at all; lowering it does the reverse. Revisit it when a defect it
    /// exposes gets fixed, since each fix removes a hurdle.
    /// </para>
    /// </summary>
    private const int HostileOdds = 8;

    /// <summary>
    /// Whether this particular field should take its hostile shape.
    ///
    /// The comparison is against the <em>last</em> value in the range rather than zero, and that
    /// matters more than it looks. <see cref="FuzzBytes"/> returns zero once the input is spent,
    /// so a comparison against zero makes an empty or short input the <em>most</em> hostile
    /// package the generator can build — every field broken at once. libFuzzer tries the empty
    /// input first, so the very first execution of every run produced a workbook with no loadable
    /// sheet at all. The degenerate input should describe the ordinary workbook, not the worst
    /// one.
    /// </summary>
    private static bool Hostile(FuzzBytes input)
    {
        return input.Int(0, HostileOdds - 1) == HostileOdds - 1;
    }

    public static byte[] Generate(FuzzBytes input)
    {
        var sheetCount = input.Int(1, 3);
        var sheets = new List<GeneratedSheet>(sheetCount);
        for (var i = 0; i < sheetCount; i++)
            sheets.Add(GenerateSheet(input, i));

        var buffer = new MemoryStream();
        using (var archive = new ZipArchive(buffer, ZipArchiveMode.Create, leaveOpen: true))
        {
            Write(archive, "[Content_Types].xml", ContentTypes(sheets));
            Write(archive, "_rels/.rels", RootRelationships());
            Write(archive, "xl/workbook.xml", WorkbookXml(input, sheets));
            Write(archive, "xl/_rels/workbook.xml.rels", WorkbookRelationships(sheets));
            Write(archive, "xl/styles.xml", StylesXml());
            Write(archive, "xl/sharedStrings.xml", SharedStringsXml());

            foreach (var sheet in sheets)
                Write(archive, sheet.PartName, sheet.Xml);
        }

        return buffer.ToArray();
    }

    private static GeneratedSheet GenerateSheet(FuzzBytes input, int index)
    {
        var ordinal = index + 1;

        // A name the fuzzer chooses freely, or a duplicate of the first sheet's. Duplicate sheet
        // names are invalid but nothing stops a producer emitting them.
        var name = input.Bool() ? $"Sheet{ordinal}" : input.Text(12);
        if (name.Length == 0)
            name = $"Sheet{ordinal}";

        var rows = input.Int(0, 6);
        var columns = input.Int(0, 6);

        var cells = new StringBuilder();
        for (var row = 1; row <= rows; row++)
        {
            cells.Append(CultureInfo.InvariantCulture, $"<row r=\"{row}\">");
            for (var column = 1; column <= columns; column++)
                cells.Append(CellXml(input, row, column));

            cells.Append("</row>");
        }

        // The declared extent either agrees with the cells written, or does not. XLibur is
        // entitled to trust dimension as a hint; it is not entitled to fault when the hint lies.
        var dimension = input.Int(0, 3) switch
        {
            0 => $"A1:{ColumnName(Math.Max(columns, 1))}{Math.Max(rows, 1)}",
            1 => "A1:A1",
            2 => "A1:XFD1048576",
            _ => "B2:A1",
        };

        var xml =
            $"""
             <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
             <worksheet xmlns="{SpreadsheetMlNamespace}">
               <dimension ref="{Escape(dimension)}" />
               <sheetData>{cells}</sheetData>
             </worksheet>
             """;

        return new GeneratedSheet(name, ordinal, $"xl/worksheets/sheet{ordinal}.xml", xml);
    }

    private static string CellXml(FuzzBytes input, int row, int column)
    {
        var reference = $"{ColumnName(column)}{row}";

        // A style index that exists, or one past the end of the cell-format table. Rare per cell,
        // because a sheet can hold dozens of cells and one bad index per package is the
        // interesting case — not every cell being bad, which only ever finds the first one.
        var styleIndex = Hostile(input) ? input.Int(CellFormatCount, CellFormatCount + 40) : input.Int(0, CellFormatCount - 1);
        var style = $" s=\"{styleIndex.ToString(CultureInfo.InvariantCulture)}\"";

        return input.Int(0, 5) switch
        {
            // A number, including the shapes that overflow a double when parsed naively.
            0 => $"<c r=\"{reference}\"{style}><v>{Escape(input.Pick("0", "1", "-1", "1e308", "1e309", "0.1", "-0", "1E-320"))}</v></c>",

            // An inline string.
            1 => $"<c r=\"{reference}\"{style} t=\"inlineStr\"><is><t>{Escape(input.Text(8))}</t></is></c>",

            // A shared string, by an index that exists or one that does not.
            2 => $"<c r=\"{reference}\"{style} t=\"s\"><v>{(Hostile(input) ? input.Int(SharedStringCount, SharedStringCount + 40) : input.Int(0, SharedStringCount - 1))}</v></c>",

            // A boolean, including values outside the two Excel defines.
            3 => $"<c r=\"{reference}\"{style} t=\"b\"><v>{input.Pick("0", "1", "2", "-1")}</v></c>",

            // An error literal, including one Excel does not define.
            4 => $"<c r=\"{reference}\"{style} t=\"e\"><v>{Escape(input.Pick("#DIV/0!", "#REF!", "#N/A", "#NOPE!"))}</v></c>",

            // A formula, including references off the end of the grid and to another sheet.
            _ => $"<c r=\"{reference}\"{style}><f>{Escape(input.Pick("1+1", "A1", "SUM(A1:B2)", "XFD1048576", "'no such sheet'!A1", "A1:A1048577"))}</f></c>",
        };
    }

    private static string WorkbookXml(FuzzBytes input, List<GeneratedSheet> sheets)
    {
        var entries = new StringBuilder();
        foreach (var sheet in sheets)
        {
            // The relationship id either names the sheet's own part, or names nothing at all.
            // A dangling r:id is the single most likely shape to be mishandled, because the
            // happy path never produces one.
            //
            // The first sheet always keeps a valid one. A workbook where *every* sheet is
            // unloadable cannot be saved at all (D31, open), so generating one ends the run on a
            // defect that is already recorded and hides everything behind it. Guaranteeing one
            // loadable sheet keeps dangling ids in play while leaving the workbook saveable.
            var mayDangle = sheet.Ordinal > 1;
            var relationshipId = mayDangle && Hostile(input)
                ? $"rIdMissing{input.Int(0, 9)}"
                : sheet.RelationshipId;

            // Sheet ids that collide, and ids outside the range Excel writes.
            var sheetId = Hostile(input) ? input.Pick(0, 1, int.MaxValue) : sheet.Ordinal;

            entries.Append(
                CultureInfo.InvariantCulture,
                $"<sheet name=\"{Escape(sheet.Name)}\" sheetId=\"{sheetId}\" r:id=\"{Escape(relationshipId)}\" />");
        }

        return
            $"""
             <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
             <workbook xmlns="{SpreadsheetMlNamespace}" xmlns:r="{RelationshipsNamespace}">
               <sheets>{entries}</sheets>
             </workbook>
             """;
    }

    private static string WorkbookRelationships(List<GeneratedSheet> sheets)
    {
        var entries = new StringBuilder();
        foreach (var sheet in sheets)
        {
            entries.Append(
                CultureInfo.InvariantCulture,
                $"""<Relationship Id="{sheet.RelationshipId}" Type="{RelationshipsNamespace}/worksheet" Target="worksheets/sheet{sheet.Ordinal}.xml" />""");
        }

        entries.Append(
            CultureInfo.InvariantCulture,
            $"""<Relationship Id="rIdStyles" Type="{RelationshipsNamespace}/styles" Target="styles.xml" />""");
        entries.Append(
            CultureInfo.InvariantCulture,
            $"""<Relationship Id="rIdStrings" Type="{RelationshipsNamespace}/sharedStrings" Target="sharedStrings.xml" />""");

        return
            $"""
             <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
             <Relationships xmlns="{PackageRelationshipsNamespace}">{entries}</Relationships>
             """;
    }

    private static string RootRelationships()
    {
        return
            $"""
             <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
             <Relationships xmlns="{PackageRelationshipsNamespace}">
               <Relationship Id="rId1" Type="{RelationshipsNamespace}/officeDocument" Target="xl/workbook.xml" />
             </Relationships>
             """;
    }

    private static string ContentTypes(List<GeneratedSheet> sheets)
    {
        var overrides = new StringBuilder();
        foreach (var sheet in sheets)
        {
            overrides.Append(
                CultureInfo.InvariantCulture,
                $"""<Override PartName="/{sheet.PartName}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" />""");
        }

        return
            $"""
             <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
             <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
               <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />
               <Default Extension="xml" ContentType="application/xml" />
               <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" />
               <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml" />
               <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml" />
               {overrides}
             </Types>
             """;
    }

    private static string StylesXml()
    {
        return
            $"""
             <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
             <styleSheet xmlns="{SpreadsheetMlNamespace}">
               <fonts count="1"><font><sz val="11" /><name val="Calibri" /></font></fonts>
               <fills count="1"><fill><patternFill patternType="none" /></fill></fills>
               <borders count="1"><border /></borders>
               <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" /></cellStyleXfs>
               <cellXfs count="{CellFormatCount}">
                 <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" />
                 <xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" />
               </cellXfs>
             </styleSheet>
             """;
    }

    private static string SharedStringsXml()
    {
        return
            $"""
             <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
             <sst xmlns="{SpreadsheetMlNamespace}" count="{SharedStringCount}" uniqueCount="{SharedStringCount}">
               <si><t>alpha</t></si>
               <si><t/></si>
               <si><t xml:space="preserve"> trailing </t></si>
             </sst>
             """;
    }

    private static void Write(ZipArchive archive, string name, string content)
    {
        using var stream = archive.CreateEntry(name).Open();
        var bytes = Encoding.UTF8.GetBytes(content);
        stream.Write(bytes, 0, bytes.Length);
    }

    private static string ColumnName(int column)
    {
        var name = string.Empty;
        while (column > 0)
        {
            var remainder = (column - 1) % 26;
            name = (char)('A' + remainder) + name;
            column = (column - 1) / 26;
        }

        return name.Length == 0 ? "A" : name;
    }

    /// <summary>
    /// Escape for an XML text node or attribute value. The generated XML must always parse —
    /// a package that fails at the XML layer tests the XML reader, which is not the point.
    /// </summary>
    private static string Escape(string value)
    {
        return value
            .Replace("&", "&amp;", StringComparison.Ordinal)
            .Replace("<", "&lt;", StringComparison.Ordinal)
            .Replace(">", "&gt;", StringComparison.Ordinal)
            .Replace("\"", "&quot;", StringComparison.Ordinal)
            .Replace("'", "&apos;", StringComparison.Ordinal);
    }

    private sealed record GeneratedSheet(string Name, int Ordinal, string PartName, string Xml)
    {
        public string RelationshipId => $"rIdSheet{Ordinal}";
    }
}
