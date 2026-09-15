
namespace XLibur.Excel;

public abstract class XLValidationCriteria : IXLValidationCriteria
{
    protected IXLDataValidation dataValidation;

    protected XLValidationCriteria(IXLDataValidation dataValidation)
    {
        this.dataValidation = dataValidation;
    }

    #region IXLValidationCriteria Members

    public void Between(string minValue, string maxValue)
    {
        dataValidation.MinValue = minValue;
        dataValidation.MaxValue = maxValue;
        dataValidation.Operator = XLOperator.Between;
    }

    public void Between(IXLCell minValue, IXLCell maxValue)
    {
        dataValidation.MinValue = CellReference(minValue);
        dataValidation.MaxValue = CellReference(maxValue);
        dataValidation.Operator = XLOperator.Between;
    }

    public void EqualOrGreaterThan(string value)
    {
        dataValidation.Value = value;
        dataValidation.Operator = XLOperator.EqualOrGreaterThan;
    }

    public void EqualOrGreaterThan(IXLCell cell)
    {
        dataValidation.Value = CellReference(cell);
        dataValidation.Operator = XLOperator.EqualOrGreaterThan;
    }

    public void EqualOrLessThan(string value)
    {
        dataValidation.Value = value;
        dataValidation.Operator = XLOperator.EqualOrLessThan;
    }

    public void EqualOrLessThan(IXLCell cell)
    {
        dataValidation.Value = CellReference(cell);
        dataValidation.Operator = XLOperator.EqualOrLessThan;
    }

    public void EqualTo(string value)
    {
        dataValidation.Value = value;
        dataValidation.Operator = XLOperator.EqualTo;
    }

    public void EqualTo(IXLCell cell)
    {
        dataValidation.Value = CellReference(cell);
        dataValidation.Operator = XLOperator.EqualTo;
    }

    public void GreaterThan(string value)
    {
        dataValidation.Value = value;
        dataValidation.Operator = XLOperator.GreaterThan;
    }

    public void GreaterThan(IXLCell cell)
    {
        dataValidation.Value = CellReference(cell);
        dataValidation.Operator = XLOperator.GreaterThan;
    }

    public void LessThan(string value)
    {
        dataValidation.Value = value;
        dataValidation.Operator = XLOperator.LessThan;
    }

    public void LessThan(IXLCell cell)
    {
        dataValidation.Value = CellReference(cell);
        dataValidation.Operator = XLOperator.LessThan;
    }

    public void NotBetween(string minValue, string maxValue)
    {
        dataValidation.MinValue = minValue;
        dataValidation.MaxValue = maxValue;
        dataValidation.Operator = XLOperator.NotBetween;
    }

    public void NotBetween(IXLCell minValue, IXLCell maxValue)
    {
        dataValidation.MinValue = CellReference(minValue);
        dataValidation.MaxValue = CellReference(maxValue);
        dataValidation.Operator = XLOperator.NotBetween;
    }

    public void NotEqualTo(string value)
    {
        dataValidation.Value = value;
        dataValidation.Operator = XLOperator.NotEqualTo;
    }

    public void NotEqualTo(IXLCell cell)
    {
        dataValidation.Value = CellReference(cell);
        dataValidation.Operator = XLOperator.NotEqualTo;
    }

    #endregion IXLValidationCriteria Members

    /// <summary>
    /// The criterion text for <paramref name="cell"/>: its fixed A1 address, with the name of its sheet,
    /// quoted where the name needs it, when that is not the rule's own sheet.
    /// </summary>
    /// <remarks>
    /// <para>
    /// A cell on another sheet needs the name. Without it, the address reads as a cell on the rule's own
    /// sheet (#524). With it, a save writes the rule in the <c>x14</c> extension, where Excel writes a
    /// rule that refers to another sheet, and a rename or a delete of that sheet reaches the criterion
    /// (D63).
    /// </para>
    /// <para>
    /// A cell on the rule's own sheet is stored without the name, as before, and as Excel writes such a
    /// reference even when it was typed with one. As nothing names the sheet, the criterion still points at
    /// its own sheet's cells after that sheet is renamed, and on a copy of the sheet at the copy's cells.
    /// A rule XLibur did not create has no sheet to compare with, so its criterion always gets the name,
    /// which is never wrong.
    /// </para>
    /// <para>
    /// A formula in a file is A1 text whatever <see cref="IXLWorkbook.ReferenceStyle"/> says, so the
    /// address is always A1.
    /// </para>
    /// </remarks>
    private string CellReference(IXLCell cell)
    {
        var onRuleSheet = dataValidation is XLDataValidation rule && ReferenceEquals(cell.Worksheet, rule.Worksheet);
        return cell.Address.ToStringFixed(XLReferenceStyle.A1, includeSheet: !onRuleSheet);
    }
}
