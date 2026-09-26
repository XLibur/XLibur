using System;

namespace XLibur.Excel;

public class XLDateCriteria : XLValidationCriteria
{
    public XLDateCriteria(IXLDataValidation dataValidation)
        : base(dataValidation)
    {
    }

    public void Between(DateTime minValue, DateTime maxValue) => base.Between(GetXLDate(minValue), GetXLDate(maxValue));

    public void EqualOrGreaterThan(DateTime value) => base.EqualOrGreaterThan(GetXLDate(value));

    public void EqualOrLessThan(DateTime value) => base.EqualOrLessThan(GetXLDate(value));

    public void EqualTo(DateTime value) => base.EqualTo(GetXLDate(value));

    public void GreaterThan(DateTime value) => base.GreaterThan(GetXLDate(value));

    public void LessThan(DateTime value) => base.LessThan(GetXLDate(value));

    public void NotBetween(DateTime minValue, DateTime maxValue) => base.NotBetween(GetXLDate(minValue), GetXLDate(maxValue));

    public void NotEqualTo(DateTime value) => base.NotEqualTo(GetXLDate(value));

    private static string GetXLDate(DateTime value) => value.ToOADate().ToInvariantString();
}
