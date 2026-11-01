using System.IO;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using XLibur.Extensions;

namespace XLibur.Tests.Extensions;

public class OpenXmlPartContainerExtensionsTests
{
    [Test]
    public async Task GetPartOrNull_ResolvesKnownId_AndReturnsNullOtherwise()
    {
        using var ms = new MemoryStream();
        using var document = SpreadsheetDocument.Create(ms, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>("rId7");

        await Assert.That(workbookPart.GetPartOrNull<WorksheetPart>("rId7")).IsSameReferenceAs(worksheetPart);
        await Assert.That(workbookPart.GetPartOrNull<OpenXmlPart>("rId7")).IsSameReferenceAs(worksheetPart);

        // Unknown id, wrong part type, and no id at all.
        await Assert.That(workbookPart.GetPartOrNull<WorksheetPart>("rId8")).IsNull();
        await Assert.That(workbookPart.GetPartOrNull<SharedStringTablePart>("rId7")).IsNull();
        await Assert.That(workbookPart.GetPartOrNull<WorksheetPart>(null)).IsNull();
        await Assert.That(workbookPart.GetPartOrNull<WorksheetPart>(string.Empty)).IsNull();
    }

    [Test]
    public async Task HasPartWithId_FindsKnownId_AndIsFalseOtherwise()
    {
        using var ms = new MemoryStream();
        using var document = SpreadsheetDocument.Create(ms, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.AddNewPart<WorksheetPart>("rId7");

        await Assert.That(workbookPart.HasPartWithId("rId7")).IsTrue();
        await Assert.That(workbookPart.HasPartWithId("rId8")).IsFalse();

        // No id at all: false, as the old linear scan answered.
        await Assert.That(workbookPart.HasPartWithId(null)).IsFalse();
        await Assert.That(workbookPart.HasPartWithId(string.Empty)).IsFalse();
    }
}
