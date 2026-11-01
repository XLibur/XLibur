using System.Collections.Generic;
using System.Data;
using System.Diagnostics.CodeAnalysis;

namespace XLibur.Excel;

public interface IXLWorksheets : IEnumerable<IXLWorksheet>
{
    int Count { get; }

    IXLWorksheet Add();

    IXLWorksheet Add(int position);

    IXLWorksheet Add(string sheetName);

    IXLWorksheet Add(string sheetName, int position);

    IXLWorksheet Add(DataTable dataTable);

    IXLWorksheet Add(DataTable dataTable, string sheetName);

    IXLWorksheet Add(DataTable dataTable, string sheetName, string tableName);

    void Add(DataSet dataSet);

    bool Contains(string sheetName);

    void Delete(string sheetName);

    void Delete(int position);

    bool TryGetWorksheet(string sheetName, [NotNullWhen(true)] out IXLWorksheet? worksheet);

    IXLWorksheet Worksheet(string sheetName);

    IXLWorksheet Worksheet(int position);
}
