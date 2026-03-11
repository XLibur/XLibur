using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;
using static XLibur.Excel.XLWorkbook;

namespace XLibur.Excel.IO;

internal class CalculationChainPartWriter
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
            foreach (var c in worksheet.Internals.CellsCollection.GetCells().Where(c => c.HasFormula))
            {
                if (c.Formula!.Type == FormulaType.DataTable)
                {
                    // Do nothing, Excel doesn't generate calc chain for data table
                }
                else if (c.HasArrayFormula)
                {
                    c.FormulaReference ??= c.AsRange().RangeAddress;

                    if (c.FormulaReference.FirstAddress.Equals(c.Address))
                    {
                        var cc = new CalculationCell
                        {
                            CellReference = c.Address.ToString(),
                            SheetId = (int)worksheet.SheetId,
                            Array = true
                        };

                        calculationChain.AppendChild(cc);

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
                else
                {
                    calculationChain.AppendChild(new CalculationCell
                    {
                        CellReference = c.Address.ToString(),
                        SheetId = (int)worksheet.SheetId
                    });
                }
            }
        }

        if (!calculationChain.Any())
            workbookPart.DeletePart(workbookPart.CalculationChainPart);
    }
}
