using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using TUnit.Assertions.Enums;
using XLibur.Excel;
using XLibur.Tests.Utils;

namespace XLibur.Tests.Excel.Saving;

/// <summary>
/// A workbook XLibur loads and saves unchanged has to load again (#743). These are the cases where
/// the save wrote a workbook part, or a package around it, that XLibur then refused to read.
/// </summary>
public class WorkbookPartReloadTests
{
    private const string SheetsWithoutRelId = @"TryToLoad\SheetsWithoutRelId.xlsx";
    private const string MacroSheetsWithoutRelId = @"TryToLoad\LO\xlsm\tdf111974.xlsm";
    private const string StrictWorkbook = @"TryToLoad\LO\xlsx\universal-content-strict.xlsx";

    [Test]
    [Arguments(SheetsWithoutRelId)]
    [Arguments(MacroSheetsWithoutRelId)]
    public async Task A_sheet_declared_with_no_relationship_id_is_declared_once_after_a_save(string resource)
    {
        // LibreOffice writes some sheets as <sheet r:id="">. The load makes each an empty worksheet,
        // as Excel does, and the save declared it a second time beside the old element.
        using var source = TestHelper.LoadWorkbook(resource);
        var sourceSheets = source.Worksheets.Select(ws => ws.Name).ToList();

        using var saved = TestHelper.LoadAndSave(resource);

        var declared = DeclaredSheets(saved);
        await Assert.That(declared.Select(s => s.Name)).IsEquivalentTo(sourceSheets, CollectionOrdering.Matching);
        await Assert.That(declared.All(s => !string.IsNullOrEmpty(s.RelId))).IsTrue();

        saved.Position = 0;
        using var reloaded = new XLWorkbook(saved);
        await Assert.That(reloaded.Worksheets.Select(ws => ws.Name)).IsEquivalentTo(sourceSheets, CollectionOrdering.Matching);
    }

    [Test]
    public async Task A_sheet_declared_with_no_relationship_id_keeps_its_place_and_visibility()
    {
        using var saved = TestHelper.LoadAndSave(MacroSheetsWithoutRelId);
        using var reloaded = new XLWorkbook(saved);

        // Positions 5 and 6 of ten, between sheets the file backs with parts.
        await Assert.That(reloaded.Worksheet("Roll weeks macro").Position).IsEqualTo(5);
        await Assert.That(reloaded.Worksheet("Roll weeks macro").Visibility).IsEqualTo(XLWorksheetVisibility.VeryHidden);
        await Assert.That(reloaded.Worksheet("Module1").Position).IsEqualTo(6);
        await Assert.That(reloaded.Worksheet("Module1").Visibility).IsEqualTo(XLWorksheetVisibility.VeryHidden);
    }

    [Test]
    public async Task A_sheet_declared_with_no_relationship_id_leaves_the_sheet_scope_indexes_of_its_neighbours_right()
    {
        // localSheetId counts every <sheet>, so a stray element shifted the scope of the names after it.
        using var saved = TestHelper.LoadAndSave(MacroSheetsWithoutRelId);

        var printAreaScopes = DeclaredDefinedNames(saved)
            .Where(n => n.Name == "_xlnm.Print_Area")
            .ToDictionary(n => n.Text, n => n.LocalSheetId);
        await Assert.That(printAreaScopes["'OP PLAN'!$A$2:$CA$123"]).IsEqualTo(1u);
        await Assert.That(printAreaScopes["'OP LTR'!$A$1:$C$56"]).IsEqualTo(8u);

        saved.Position = 0;
        using var reloaded = new XLWorkbook(saved);
        var printArea = reloaded.Worksheet("OP LTR").PageSetup.PrintAreas.Single();
        await Assert.That(printArea.RangeAddress.ToStringRelative()).IsEqualTo("A1:C56");
    }

    [Test]
    public async Task A_deleted_sheet_that_was_declared_with_no_relationship_id_is_not_declared_after_a_save()
    {
        // The sheet has no relationship id to record as deleted, so the save kept its <sheet>.
        using var wb = TestHelper.LoadWorkbook(SheetsWithoutRelId);
        wb.Worksheet("Sheet without relId").Delete();

        using var reloaded = TestHelper.SaveAndReload(wb);

        await Assert.That(reloaded.Worksheets.Select(ws => ws.Name))
            .IsEquivalentTo(["Sheet1", "Pivot Sheet without relId"], CollectionOrdering.Matching);
    }

    [Test]
    public async Task A_strict_workbook_saves_a_package_with_no_relationship_to_a_part_it_deleted()
    {
        // The source has a calcChain but no formulas, so the save deletes the calcChain part. The SDK
        // put the Strict relationship to it back when it closed the package.
        using var saved = TestHelper.LoadAndSave(StrictWorkbook);

        await Assert.That(DanglingRelationships(saved)).IsEmpty();

        saved.Position = 0;
        using var reloaded = new XLWorkbook(saved);
        await Assert.That(reloaded.Worksheets.Select(ws => ws.Name))
            .IsEquivalentTo(["Sheet1", "Sheet2"], CollectionOrdering.Matching);
    }

    [Test]
    public async Task A_strict_workbook_saves_a_package_with_no_relationship_to_a_sheet_deleted_from_it()
    {
        // Not only the calcChain: any part the save deletes from a Strict package was affected.
        using var wb = TestHelper.LoadWorkbook(StrictWorkbook);
        wb.Worksheet("Sheet2").Delete();

        using var saved = new MemoryStream();
        wb.SaveAs(saved);

        await Assert.That(DanglingRelationships(saved)).IsEmpty();

        saved.Position = 0;
        using var reloaded = new XLWorkbook(saved);
        await Assert.That(reloaded.Worksheets.Select(ws => ws.Name)).IsEquivalentTo(["Sheet1"]);
    }

    [Test]
    [Arguments(false)]
    [Arguments(true)]
    public async Task A_strict_workbook_loaded_from_a_file_saves_a_file_that_loads_again(bool saveAsAnotherFile)
    {
        // A workbook loaded from a path is saved through the overload that opens the package by path,
        // whether Save writes the file it came from or SaveAs a copy of it.
        using var source = new TemporaryFile();
        using var copy = new TemporaryFile();
        using (var resource = TestHelper.OpenResource(StrictWorkbook))
            await File.WriteAllBytesAsync(source.Path, resource.ToArray());

        var target = saveAsAnotherFile ? copy.Path : source.Path;
        using (var wb = new XLWorkbook(source.Path))
        {
            if (saveAsAnotherFile)
                wb.SaveAs(target);
            else
                wb.Save();
        }

        using (var saved = new MemoryStream(await File.ReadAllBytesAsync(target)))
            await Assert.That(DanglingRelationships(saved)).IsEmpty();

        using var reloaded = new XLWorkbook(target);
        await Assert.That(reloaded.Worksheets.Count).IsEqualTo(2);
    }

    [Test]
    public async Task A_strict_workbook_is_saved_with_transitional_relationships()
    {
        using var saved = TestHelper.LoadAndSave(StrictWorkbook);

        using var package = SpreadsheetDocument.Open(saved, false);
        await Assert.That(package.StrictRelationshipFound).IsFalse();
    }

    private static List<(string Name, string? RelId)> DeclaredSheets(MemoryStream saved)
    {
        saved.Position = 0;
        using var package = SpreadsheetDocument.Open(saved, false);
        return package.WorkbookPart!.Workbook!.Sheets!.Elements<Sheet>()
            .Select(s => (s.Name!.Value!, s.Id?.Value))
            .ToList();
    }

    private static List<(string Name, uint? LocalSheetId, string Text)> DeclaredDefinedNames(MemoryStream saved)
    {
        saved.Position = 0;
        using var package = SpreadsheetDocument.Open(saved, false);
        return package.WorkbookPart!.Workbook!.DefinedNames!.Elements<DefinedName>()
            .Select(n => (n.Name!.Value!, n.LocalSheetId?.Value, n.Text))
            .ToList();
    }

    /// <summary>
    /// Every internal relationship in the package, from the package itself or from one of its parts,
    /// whose target part the package does not hold.
    /// </summary>
    private static List<string> DanglingRelationships(MemoryStream saved)
    {
        saved.Position = 0;
        using var package = Package.Open(saved, FileMode.Open, FileAccess.Read);
        var dangling = new List<string>();

        void Check(Uri sourceUri, IEnumerable<PackageRelationship> relationships)
        {
            foreach (var relationship in relationships.Where(r => r.TargetMode == TargetMode.Internal))
            {
                var target = PackUriHelper.ResolvePartUri(sourceUri, relationship.TargetUri);
                if (!package.PartExists(target))
                    dangling.Add($"{sourceUri} {relationship.Id} -> {target}");
            }
        }

        Check(new Uri("/", UriKind.Relative), package.GetRelationships());
        foreach (var part in package.GetParts().Where(p => !p.Uri.OriginalString.EndsWith(".rels", StringComparison.Ordinal)))
            Check(part.Uri, part.GetRelationships());

        return dangling;
    }
}
