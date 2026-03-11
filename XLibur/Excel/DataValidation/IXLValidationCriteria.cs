
namespace ClosedXML.Excel;

public interface IXLValidationCriteria
{
    void Between(string minValue, string maxValue);

    void Between(IXLCell minValue, IXLCell maxValue);

    void EqualOrGreaterThan(string value);

    void EqualOrGreaterThan(IXLCell cell);

    void EqualOrLessThan(string value);

    void EqualOrLessThan(IXLCell cell);

    void EqualTo(string value);

    void EqualTo(IXLCell cell);

    void GreaterThan(string value);

    void GreaterThan(IXLCell cell);

    void LessThan(string value);

    void LessThan(IXLCell cell);

    void NotBetween(string minValue, string maxValue);

    void NotBetween(IXLCell minValue, IXLCell maxValue);

    void NotEqualTo(string value);

    void NotEqualTo(IXLCell cell);
}
