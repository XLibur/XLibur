using XLibur.Examples;
using XLibur.Examples.Misc;
using System.Threading.Tasks;

namespace XLibur.Tests.Examples;

public class MiscTests
{
    [Test]
    public async Task AddingDataSet()
    {
        await TestHelper.RunTestExample<AddingDataSet>(@"Misc\AddingDataSet.xlsx");
    }

    [Test]
    public async Task AddingDataTableAsWorksheet()
    {
        await TestHelper.RunTestExample<AddingDataTableAsWorksheet>(@"Misc\AddingDataTableAsWorksheet.xlsx");
    }

    [Test]
    // Windows-only: AdjustToContents produces font-dependent column widths; reference xlsx was generated on Windows with Calibri
    [RunOn(TUnit.Core.Enums.OS.Windows)]
    public async Task AdjustToContents()
    {
        await TestHelper.RunTestExample<AdjustToContents>(@"Misc\AdjustToContents.xlsx");
    }

    [Test]
    public async Task AdjustToContentsWithAutoFilter()
    {
        await TestHelper.RunTestExample<AdjustToContentsWithAutoFilter>(@"Misc\AdjustToContentsWithAutoFilter.xlsx");
    }

    [Test]
    public async Task AutoFilter()
    {
        await TestHelper.RunTestExample<AutoFilter>(@"Misc\AutoFilter.xlsx");
    }

    [Test]
    public async Task BasicTable()
    {
        await TestHelper.RunTestExample<BasicTable>(@"Misc\BasicTable.xlsx");
    }

    [Test]
    public async Task BlankCells()
    {
        await TestHelper.RunTestExample<BlankCells>(@"Misc\BlankCells.xlsx");
    }

    [Test]
    // Windows-only: SharedStrings differ on Linux due to platform-specific double-to-string formatting
    [RunOn(TUnit.Core.Enums.OS.Windows)]
    public async Task CellValues()
    {
        await TestHelper.RunTestExample<CellValues>(@"Misc\CellValues.xlsx", true);
    }

    [Test]
    public async Task Collections()
    {
        await TestHelper.RunTestExample<Collections>(@"Misc\Collections.xlsx");
    }

    [Test]
    public async Task CopyingRowsAndColumns()
    {
        await TestHelper.RunTestExample<CopyingRowsAndColumns>(@"Misc\CopyingRowsAndColumns.xlsx");
    }

    [Test]
    public async Task CopyingWorksheets()
    {
        await TestHelper.RunTestExample<CopyingWorksheets>(@"Misc\CopyingWorksheets.xlsx");
    }

    [Test]
    // Windows-only: AdjustToContents produces font-dependent column widths; reference xlsx was generated on Windows with Calibri
    [RunOn(TUnit.Core.Enums.OS.Windows)]
    public async Task DataTypes()
    {
        await TestHelper.RunTestExample<DataTypes>(@"Misc\DataTypes.xlsx");
    }

    [Test]
    public async Task DataValidation()
    {
        await TestHelper.RunTestExample<DataValidation>(@"Misc\DataValidation.xlsx");
    }

    [Test]
    public async Task DataValidationDecimal()
    {
        await TestHelper.RunTestExample<DataValidationDecimal>(@"Misc\DataValidationDecimal.xlsx");
    }

    [Test]
    public async Task DataValidationWholeNumber()
    {
        await TestHelper.RunTestExample<DataValidationWholeNumber>(@"Misc\DataValidationWholeNumber.xlsx");
    }

    [Test]
    public async Task DataValidationTextLength()
    {
        await TestHelper.RunTestExample<DataValidationTextLength>(@"Misc\DataValidationTextLength.xlsx");
    }

    [Test]
    public async Task DataValidationDate()
    {
        await TestHelper.RunTestExample<DataValidationDate>(@"Misc\DataValidationDate.xlsx");
    }

    [Test]
    public async Task DataValidationTime()
    {
        await TestHelper.RunTestExample<DataValidationTime>(@"Misc\DataValidationTime.xlsx");
    }

    [Test]
    public async Task Formulas()
    {
        await TestHelper.RunTestExample<Formulas>(@"Misc\Formulas.xlsx");
    }

    [Test]
    public async Task FormulasWithEvaluation()
    {
        await TestHelper.RunTestExample<FormulasWithEvaluation>(@"Misc\FormulasWithEvaluation.xlsx", true);
    }

    [Test]
    public async Task FreezePanes()
    {
        await TestHelper.RunTestExample<FreezePanes>(@"Misc\FreezePanes.xlsx");
    }

    [Test]
    public async Task HideSheets()
    {
        await TestHelper.RunTestExample<HideSheets>(@"Misc\HideSheets.xlsx");
    }

    [Test]
    public async Task HideUnhide()
    {
        await TestHelper.RunTestExample<HideUnhide>(@"Misc\HideUnhide.xlsx");
    }

    [Test]
    public async Task Hyperlinks()
    {
        await TestHelper.RunTestExample<Hyperlinks>(@"Misc\Hyperlinks.xlsx");
    }

    [Test]
    public async Task InsertingData()
    {
        await TestHelper.RunTestExample<InsertingData>(@"Misc\InsertingData.xlsx");
    }

    [Test]
    public async Task LambdaExpressions()
    {
        await TestHelper.RunTestExample<LambdaExpressions>(@"Misc\LambdaExpressions.xlsx");
    }

    [Test]
    public async Task MergeCells()
    {
        await TestHelper.RunTestExample<MergeCells>(@"Misc\MergeCells.xlsx");
    }

    [Test]
    public async Task MergeMoves()
    {
        await TestHelper.RunTestExample<MergeMoves>(@"Misc\MergeMoves.xlsx");
    }

    [Test]
    public async Task Outline()
    {
        await TestHelper.RunTestExample<Outline>(@"Misc\Outline.xlsx");
    }

    [Test]
    public async Task RightToLeft()
    {
        await TestHelper.RunTestExample<RightToLeft>(@"Misc\RightToLeft.xlsx");
    }

    [Test]
    public async Task SheetProtection()
    {
        await TestHelper.RunTestExample<SheetProtection>(@"Misc\SheetProtection.xlsx");
    }

    [Test]
    public async Task SheetViews()
    {
        await TestHelper.RunTestExample<SheetViews>(@"Misc\SheetViews.xlsx");
    }

    [Test]
    public async Task ShiftingFormulas()
    {
        await TestHelper.RunTestExample<ShiftingFormulas>(@"Misc\ShiftingFormulas.xlsx");
    }

    [Test]
    public async Task ShowCase()
    {
        await TestHelper.RunTestExample<ShowCase>(@"Misc\ShowCase.xlsx");
    }

    [Test]
    public async Task TabColors()
    {
        await TestHelper.RunTestExample<TabColors>(@"Misc\TabColors.xlsx");
    }

    [Test]
    public async Task WorkbookProperties()
    {
        await TestHelper.RunTestExample<WorkbookProperties>(@"Misc\WorkbookProperties.xlsx");
    }

    [Test]
    public async Task WorkbookProtection()
    {
        await TestHelper.RunTestExample<WorkbookProtection>(@"Misc\WorkbookProtection.xlsx");
    }
}
