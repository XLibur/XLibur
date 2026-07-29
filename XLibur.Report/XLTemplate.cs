using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using XLibur.Excel;
using XLibur.Report.Expressions;
using XLibur.Report.Ranges;

namespace XLibur.Report;

/// <summary>
/// An Excel workbook used as a report template.
/// </summary>
/// <example>
/// <code>
/// using var template = new XLTemplate("SalesReport.xlsx");
/// template.AddVariable("Company", "Contoso");
/// template.AddVariable("Items", sales);
/// var result = template.Generate();
/// template.SaveAs("SalesReport-2026.xlsx");
/// </code>
/// </example>
public sealed class XLTemplate : IXLTemplate
{
    private readonly Dictionary<string, object?> _variables = new(StringComparer.Ordinal);
    private readonly TemplateErrors _errors = new();
    private readonly IExpressionEngine _engine;
    private readonly bool _ownsWorkbook;
    private bool _disposed;

    /// <summary>Opens the template workbook at <paramref name="path"/>.</summary>
    public XLTemplate(string path, IExpressionEngine? engine = null)
        : this(new XLWorkbook(path), engine, ownsWorkbook: true)
    {
    }

    /// <summary>Opens the template workbook from <paramref name="stream"/>.</summary>
    public XLTemplate(Stream stream, IExpressionEngine? engine = null)
        : this(new XLWorkbook(stream), engine, ownsWorkbook: true)
    {
    }

    /// <summary>
    /// Uses an already-open <paramref name="workbook"/> as the template. The caller keeps
    /// ownership: disposing the template leaves the workbook open.
    /// </summary>
    public XLTemplate(IXLWorkbook workbook, IExpressionEngine? engine = null)
        : this(workbook, engine, ownsWorkbook: false)
    {
    }

    private XLTemplate(IXLWorkbook workbook, IExpressionEngine? engine, bool ownsWorkbook)
    {
        Workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
        _engine = engine ?? new ScribanExpressionEngine();
        _ownsWorkbook = ownsWorkbook;
    }

    /// <inheritdoc />
    public IXLWorkbook Workbook { get; }

    /// <inheritdoc />
    public bool IsGenerated { get; private set; }

    /// <summary>The engine evaluating this template's expressions.</summary>
    public IExpressionEngine Engine => _engine;

    /// <inheritdoc />
    public void AddVariable(string alias, object? value)
    {
        if (string.IsNullOrWhiteSpace(alias))
        {
            throw new ArgumentException("A variable name is required.", nameof(alias));
        }

        ThrowIfDisposed();

        _variables[alias] = value is DataTable table ? ToRowDictionaries(table) : value;
    }

    /// <summary>
    /// Binds a <see cref="DataTable"/> as a materialised list of column-keyed dictionaries.
    /// </summary>
    /// <remarks>
    /// Two reasons not to hand out <see cref="DataRow"/>s directly. A row exposes its columns
    /// through an indexer rather than through properties, so <c>{{ row.Customer }}</c> — the way a
    /// template author naturally writes it — would find nothing; and the rows have to be
    /// materialised anyway, because expanding a range enumerates its data more than once.
    /// </remarks>
    private static List<Dictionary<string, object?>> ToRowDictionaries(DataTable table)
    {
        var columns = table.Columns.Cast<DataColumn>().ToList();
        var rows = new List<Dictionary<string, object?>>(table.Rows.Count);

        foreach (DataRow row in table.Rows)
        {
            var values = new Dictionary<string, object?>(columns.Count, StringComparer.Ordinal);
            foreach (var column in columns)
            {
                var value = row[column];
                values[column.ColumnName] = value is DBNull ? null : value;
            }

            rows.Add(values);
        }

        return rows;
    }

    /// <inheritdoc />
    public void AddVariable(object value)
    {
        if (value is null)
        {
            throw new ArgumentNullException(nameof(value));
        }

        ThrowIfDisposed();

        if (value is IDictionary dictionary)
        {
            foreach (DictionaryEntry entry in dictionary)
            {
                var key = entry.Key?.ToString();
                if (!string.IsNullOrWhiteSpace(key))
                {
                    AddVariable(key!, entry.Value);
                }
            }

            return;
        }

        var type = value.GetType();

        foreach (var property in type.GetProperties(BindingFlags.Public | BindingFlags.Instance))
        {
            if (property.CanRead && property.GetIndexParameters().Length == 0)
            {
                AddVariable(property.Name, property.GetValue(value));
            }
        }

        foreach (var field in type.GetFields(BindingFlags.Public | BindingFlags.Instance))
        {
            AddVariable(field.Name, field.GetValue(value));
        }
    }

    /// <inheritdoc />
    public XLGenerateResult Generate()
    {
        ThrowIfDisposed();

        if (IsGenerated)
        {
            throw new InvalidOperationException("This template has already been generated.");
        }

        var interpreter = new RangeInterpreter(_engine, _errors);
        interpreter.Evaluate(Workbook, _variables);

        IsGenerated = true;
        return new XLGenerateResult(_errors);
    }

    /// <inheritdoc />
    public void SaveAs(string file)
    {
        ThrowIfDisposed();
        Workbook.SaveAs(file);
    }

    /// <inheritdoc />
    public void SaveAs(Stream stream)
    {
        ThrowIfDisposed();
        Workbook.SaveAs(stream);
    }

    /// <inheritdoc />
    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;

        if (_ownsWorkbook)
        {
            Workbook.Dispose();
        }
    }

    private void ThrowIfDisposed()
    {
        if (_disposed)
        {
            throw new ObjectDisposedException(nameof(XLTemplate));
        }
    }
}
