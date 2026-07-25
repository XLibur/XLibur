using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel.Coordinates;
using static XLibur.Excel.XLWorkbook;

namespace XLibur.Excel.IO;

internal static class CalculationChainPartWriter
{
    internal static void GenerateContent(WorkbookPart workbookPart, XLWorkbook workbook, SaveContext context)
    {
        if (workbookPart.CalculationChainPart == null)
            workbookPart.AddNewPart<CalculationChainPart>(context.RelIdGenerator.GetNext(RelType.Workbook));

        workbookPart.CalculationChainPart!.CalculationChain ??= new CalculationChain();

        var calculationChain = workbookPart.CalculationChainPart.CalculationChain;
        calculationChain.RemoveAllChildren<CalculationCell>();

        foreach (var worksheet in workbook.WorksheetsInternal)
        {
            // Enumerate the formula slice instead of every used cell: only formula cells need an
            // XLCell wrapper here, and on a large sheet they are a tiny fraction of the total.
            var cellsCollection = worksheet.Internals.CellsCollection;
            var formulas = cellsCollection.FormulaSlice;
            if (formulas.IsEmpty)
                continue;

            using var formulaEnumerator = formulas.GetForwardEnumerator(XLSheetRange.Full);
            while (formulaEnumerator.MoveNext())
            {
                AppendFormulaCell(calculationChain, cellsCollection.GetCell(formulaEnumerator.Point), worksheet);
            }
        }

        if (!calculationChain.Any())
            workbookPart.DeletePart(workbookPart.CalculationChainPart);
    }

    private static void AppendFormulaCell(CalculationChain calculationChain, XLCell c, XLWorksheet worksheet)
    {
        if (c.Formula!.Type == FormulaType.DataTable)
            return;

        if (c.HasArrayFormula)
        {
            AppendArrayFormulaCell(calculationChain, c, worksheet);
            return;
        }

        calculationChain.AppendChild(new CalculationCell
        {
            CellReference = c.Address.ToString(),
            SheetId = (int)worksheet.SheetId
        });
    }

    private static void AppendArrayFormulaCell(CalculationChain calculationChain, XLCell c, XLWorksheet worksheet)
    {
        c.FormulaReference ??= c.AsRange().RangeAddress;

        if (!c.FormulaReference.FirstAddress.Equals(c.Address))
            return;

        calculationChain.AppendChild(new CalculationCell
        {
            CellReference = c.Address.ToString(),
            SheetId = (int)worksheet.SheetId,
            Array = true
        });

        foreach (var childCell in worksheet.Range(c.FormulaReference).Cells())
        {
            calculationChain.AppendChild(new CalculationCell
            {
                CellReference = childCell.Address.ToString(),
                SheetId = (int)worksheet.SheetId,
            });
        }
    }
}
