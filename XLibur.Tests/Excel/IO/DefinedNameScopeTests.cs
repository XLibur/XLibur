using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using XLibur.Excel;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace XLibur.Tests.Excel.IO;

/// <summary>
/// A defined name's <c>localSheetId</c> is the 0-based position of its sheet in the workbook's
/// <c>&lt;sheets&gt;</c> list, which counts every sheet, a chartsheet included (ECMA-376). It is not the
/// sheet's <c>sheetId</c>. The two differ once a sheet has been inserted before existing ones, and the
/// reader took one for the other when it placed a print area, so the print area loaded onto the sheet
/// whose <c>sheetId</c> matched (D75, #496).
/// </summary>
public class DefinedNameScopeTests
{
    private const string PrintArea = "_xlnm.Print_Area";
    private const string PrintTitles = "_xlnm.Print_Titles";
    private const string ChartsheetBook = @"Other\PivotTableReferenceFiles\ChartsheetAndPivotTable.xlsx";

    /// <summary>
    /// Tab order <c>Alpha</c>, <c>Beta</c>, <c>Gamma</c>, with sheetIds 3, 1, 2, so that no sheet's
    /// position agrees with its sheetId: read as a sheetId, each <c>localSheetId</c> names another sheet.
    /// </summary>
    private const string OutOfOrderScopes = """
        Alpha | print area - | rows to repeat - | names Local=Alpha!$B$2
        Beta | print area OFFSET(Beta!$A$1,0,0,4,2) | rows to repeat - | names -
        Gamma | print area - | rows to repeat 1:2 | names -
        """;

    private const string ChartsheetFirstScopes = """
        Data | print area OFFSET(Data!$A$1,0,0,2,2) | rows to repeat 1:1 | names OnData=Data!$B$1
        Pivot | print area - | rows to repeat - | names Local=Pivot!$A$1
        """;

    private const string ChartsheetFirstSavedNames = """
        Data|OnData = Data!$B$1
        Data|_xlnm.Print_Area = OFFSET(Data!$A$1,0,0,2,2)
        Data|_xlnm.Print_Titles = Data!1:1
        Pivot|Local = Pivot!$A$1
        """;

    [Test]
    public async Task Names_load_onto_the_sheet_at_their_position()
    {
        using var package = OutOfOrderBook();

        using var wb = new XLWorkbook(package);

        await Assert.That(Scopes(wb)).IsEqualTo(Lines(OutOfOrderScopes));
    }

    /// <summary>
    /// The writer counts a sheet's position in the <c>&lt;sheets&gt;</c> list it writes, and keeps the
    /// sheetIds as they were, so the saved file is out of order in the same way.
    /// </summary>
    [Test]
    public async Task Names_keep_their_sheet_through_a_save_and_a_reload()
    {
        using var package = OutOfOrderBook();
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook(package))
            wb.SaveAs(saved);

        await Assert.That(SavedSheets(saved)).IsEquivalentTo(["Alpha:3", "Beta:1", "Gamma:2"]);
        await Assert.That(SavedNames(saved)).IsEqualTo(Lines([
            "Alpha|Local = Alpha!$B$2",
            $"Beta|{PrintArea} = OFFSET(Beta!$A$1,0,0,4,2)",
            $"Gamma|{PrintTitles} = Gamma!1:2",
        ]));

        using var reloaded = new XLWorkbook(saved);
        await Assert.That(Scopes(reloaded)).IsEqualTo(Lines(OutOfOrderScopes));
    }

    /// <summary>
    /// A chartsheet is in the <c>&lt;sheets&gt;</c> list, so one before a worksheet moves that
    /// worksheet's position along by one, although XLibur does not model the chartsheet.
    /// </summary>
    [Test]
    public async Task A_chartsheet_before_a_worksheet_counts_in_its_position_on_load()
    {
        using var package = ChartsheetFirstBook();

        using var wb = new XLWorkbook(package);

        await Assert.That(Scopes(wb)).IsEqualTo(Lines(ChartsheetFirstScopes));
    }

    [Test]
    public async Task A_chartsheet_before_a_worksheet_counts_in_its_position_on_save()
    {
        using var package = ChartsheetFirstBook();
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook(package))
            wb.SaveAs(saved);

        await Assert.That(SavedSheets(saved)).IsEquivalentTo(["Chart:3", "Data:1", "Pivot:2"]);
        await Assert.That(SavedNames(saved)).IsEqualTo(Lines(ChartsheetFirstSavedNames));

        using var reloaded = new XLWorkbook(saved);
        await Assert.That(Scopes(reloaded)).IsEqualTo(Lines(ChartsheetFirstScopes));
    }

    /// <summary>
    /// A name scoped to the chartsheet, or to a position past the last sheet, has no worksheet to
    /// hold it. It is dropped, where it used to fail the whole load, and the names after it still
    /// resolve at their own positions.
    /// </summary>
    [Test]
    public async Task A_name_scoped_to_no_worksheet_is_dropped_and_the_load_goes_on()
    {
        using var package = ChartsheetFirstBook(withNamesScopedToNoWorksheet: true);

        using var wb = new XLWorkbook(package);

        await Assert.That(Scopes(wb)).IsEqualTo(Lines(ChartsheetFirstScopes));
        await Assert.That(wb.DefinedNames).IsEmpty();
    }

    [Test]
    public async Task Dropping_a_name_scoped_to_no_worksheet_keeps_the_others_through_a_save_and_a_reload()
    {
        using var package = ChartsheetFirstBook(withNamesScopedToNoWorksheet: true);
        using var saved = new MemoryStream();
        using (var wb = new XLWorkbook(package))
            wb.SaveAs(saved);

        await Assert.That(SavedNames(saved)).IsEqualTo(Lines(ChartsheetFirstSavedNames));

        using var reloaded = new XLWorkbook(saved);
        await Assert.That(Scopes(reloaded)).IsEqualTo(Lines(ChartsheetFirstScopes));
    }

    /// <summary>
    /// A package built through the SDK, since this is the format's rule and not Excel's behaviour: three
    /// sheets out of sheetId order, and a print area, print titles and an ordinary name, each scoped to
    /// a different sheet.
    /// </summary>
    private static MemoryStream OutOfOrderBook()
    {
        var package = new MemoryStream();
        using (var document = SpreadsheetDocument.Create(package, SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = document.AddWorkbookPart();
            var sheets = new S.Sheets();
            foreach (var (name, sheetId) in new[] { ("Alpha", 3u), ("Beta", 1u), ("Gamma", 2u) })
            {
                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new S.Worksheet(new S.SheetData());
                sheets.Append(new S.Sheet { Name = name, SheetId = sheetId, Id = workbookPart.GetIdOfPart(worksheetPart) });
            }

            workbookPart.Workbook = new S.Workbook(sheets, new S.DefinedNames(
                new S.DefinedName { Name = PrintArea, LocalSheetId = 1, Text = "OFFSET(Beta!$A$1,0,0,4,2)" },
                new S.DefinedName { Name = PrintTitles, LocalSheetId = 2, Text = "Gamma!$1:$2" },
                new S.DefinedName { Name = "Local", LocalSheetId = 0, Text = "Alpha!$B$2" }));
        }

        package.Position = 0;
        return package;
    }

    /// <summary>
    /// The Excel-authored chartsheet workbook with the chartsheet moved to the front, so the tab order
    /// is <c>Chart</c>, <c>Data</c>, <c>Pivot</c> with sheetIds 3, 1, 2, and names scoped to both
    /// worksheets.
    /// </summary>
    /// <param name="withNamesScopedToNoWorksheet">
    /// Whether the names start with one scoped to the chartsheet and one scoped to a position past the
    /// last sheet, so that the names after them have to resolve past both.
    /// </param>
    private static MemoryStream ChartsheetFirstBook(bool withNamesScopedToNoWorksheet = false)
    {
        var package = new MemoryStream();
        using (var fixture = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(ChartsheetBook)))
            fixture.CopyTo(package);

        using (var document = SpreadsheetDocument.Open(package, true))
        {
            var workbook = document.WorkbookPart!.Workbook!;
            var chart = workbook.Sheets!.Elements<S.Sheet>().Single(s => s.Name == "Chart");
            chart.Remove();
            workbook.Sheets.PrependChild(chart);

            var names = new List<S.DefinedName>();
            if (withNamesScopedToNoWorksheet)
            {
                names.Add(new S.DefinedName { Name = "OnChart", LocalSheetId = 0, Text = "Data!$A$1" });
                names.Add(new S.DefinedName { Name = "PastTheEnd", LocalSheetId = 3, Text = "Data!$A$1" });
            }

            names.Add(new S.DefinedName { Name = PrintArea, LocalSheetId = 1, Text = "OFFSET(Data!$A$1,0,0,2,2)" });
            names.Add(new S.DefinedName { Name = PrintTitles, LocalSheetId = 1, Text = "Data!$1:$1" });
            names.Add(new S.DefinedName { Name = "OnData", LocalSheetId = 1, Text = "Data!$B$1" });
            names.Add(new S.DefinedName { Name = "Local", LocalSheetId = 2, Text = "Pivot!$A$1" });
            workbook.DefinedNames = new S.DefinedNames(names);
        }

        package.Position = 0;
        return package;
    }

    /// <summary>What each worksheet holds of the names under test, one line per sheet in tab order.</summary>
    private static string Scopes(XLWorkbook wb)
        => Lines(wb.Worksheets.OrderBy(w => w.Position).Select(w =>
        {
            var printArea = ((XLPrintAreas)w.PageSetup.PrintAreas).FormulaReference ?? "-";
            var rows = w.PageSetup.FirstRowToRepeatAtTop > 0
                ? $"{w.PageSetup.FirstRowToRepeatAtTop}:{w.PageSetup.LastRowToRepeatAtTop}"
                : "-";
            var names = w.DefinedNames.Any() ? string.Join(",", w.DefinedNames.Select(n => $"{n.Name}={n.RefersTo}")) : "-";
            return $"{w.Name} | print area {printArea} | rows to repeat {rows} | names {names}";
        }));

    /// <summary>Each <c>&lt;sheet&gt;</c> of a saved package as name and sheetId, in tab order.</summary>
    private static List<string> SavedSheets(Stream package)
    {
        package.Position = 0;
        using var document = SpreadsheetDocument.Open(package, false);
        return document.WorkbookPart!.Workbook!.Sheets!.Elements<S.Sheet>()
            .Select(s => $"{s.Name!.Value}:{s.SheetId!.Value}")
            .ToList();
    }

    /// <summary>
    /// Each defined name of a saved package, keyed by the sheet at its <c>localSheetId</c>, one line
    /// each so that a failure shows both sides in full.
    /// </summary>
    private static string SavedNames(Stream package)
    {
        package.Position = 0;
        using var document = SpreadsheetDocument.Open(package, false);
        var workbook = document.WorkbookPart!.Workbook!;
        var sheets = workbook.Sheets!.Elements<S.Sheet>().ToList();
        return Lines((workbook.DefinedNames?.Elements<S.DefinedName>() ?? [])
            .Select(n =>
            {
                var scope = n.LocalSheetId?.Value is { } id ? sheets[(int)id].Name!.Value! : "(workbook)";
                return $"{scope}|{n.Name!.Value} = {n.Text}";
            })
            .Order(StringComparer.Ordinal));
    }

    private static string Lines(string text) => Lines(text.Split('\n').Select(l => l.TrimEnd('\r')));

    private static string Lines(IEnumerable<string> items) => string.Join(Environment.NewLine, items);
}
