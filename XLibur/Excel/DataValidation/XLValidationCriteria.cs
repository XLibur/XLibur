
namespace XLibur.Excel;

public abstract class XLValidationCriteria : IXLValidationCriteria
{
    protected IXLDataValidation dataValidation;

    protected XLValidationCriteria(IXLDataValidation dataValidation)
    {
        this.dataValidation = dataValidation;
    }

    #region IXLValidationCriteria Members

    public void Between(string minValue, string maxValue) => Set(minValue, maxValue, XLOperator.Between);

    public void Between(IXLCell minValue, IXLCell maxValue)
        => Set(CellReference(minValue), CellReference(maxValue), XLOperator.Between);

    public void EqualOrGreaterThan(string value) => Set(value, XLOperator.EqualOrGreaterThan);

    public void EqualOrGreaterThan(IXLCell cell) => Set(CellReference(cell), XLOperator.EqualOrGreaterThan);

    public void EqualOrLessThan(string value) => Set(value, XLOperator.EqualOrLessThan);

    public void EqualOrLessThan(IXLCell cell) => Set(CellReference(cell), XLOperator.EqualOrLessThan);

    public void EqualTo(string value) => Set(value, XLOperator.EqualTo);

    public void EqualTo(IXLCell cell) => Set(CellReference(cell), XLOperator.EqualTo);

    public void GreaterThan(string value) => Set(value, XLOperator.GreaterThan);

    public void GreaterThan(IXLCell cell) => Set(CellReference(cell), XLOperator.GreaterThan);

    public void LessThan(string value) => Set(value, XLOperator.LessThan);

    public void LessThan(IXLCell cell) => Set(CellReference(cell), XLOperator.LessThan);

    public void NotBetween(string minValue, string maxValue) => Set(minValue, maxValue, XLOperator.NotBetween);

    public void NotBetween(IXLCell minValue, IXLCell maxValue)
        => Set(CellReference(minValue), CellReference(maxValue), XLOperator.NotBetween);

    public void NotEqualTo(string value) => Set(value, XLOperator.NotEqualTo);

    public void NotEqualTo(IXLCell cell) => Set(CellReference(cell), XLOperator.NotEqualTo);

    #endregion IXLValidationCriteria Members

    /// <summary>
    /// Stores a criterion that compares with one value: the one place every criteria type writes it.
    /// </summary>
    private protected void Set(string value, XLOperator op)
    {
        dataValidation.Value = value;
        dataValidation.Operator = op;
    }

    /// <summary>
    /// Stores a criterion that compares with two bounds (between, not between).
    /// </summary>
    private protected void Set(string minValue, string maxValue, XLOperator op)
    {
        dataValidation.MinValue = minValue;
        dataValidation.MaxValue = maxValue;
        dataValidation.Operator = op;
    }

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
