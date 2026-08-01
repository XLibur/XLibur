using System;
using System.Collections.Generic;

namespace XLibur.Report.Benchmarks;

/// <summary>
/// One row of the benchmark's data: ten members of the kinds a real report binds — text, integers,
/// decimals, a date, a bool and a computed property.
/// </summary>
/// <remarks>
/// Ten columns rather than three because the per-cell costs (expression evaluation, value conversion,
/// cell writes) scale with the column count while the per-row costs (the row insert, the block copy)
/// do not, and a three-column workload would flatter the engine by hiding the former behind the latter.
/// </remarks>
public sealed class ReportRow
{
    public string Region { get; init; } = string.Empty;

    public string Category { get; init; } = string.Empty;

    public string Product { get; init; } = string.Empty;

    public string Reference { get; init; } = string.Empty;

    public int Quantity { get; init; }

    public decimal UnitPrice { get; init; }

    public decimal Discount { get; init; }

    public DateTime SoldOn { get; init; }

    public bool IsExport { get; init; }

    public decimal Total => Quantity * UnitPrice * (1m - Discount);
}

/// <summary>The benchmark's data, generated deterministically so runs are comparable.</summary>
public static class ReportData
{
    private static readonly string[] Regions = { "North", "South", "East", "West", "Central" };
    private static readonly string[] Categories = { "Retail", "Trade", "Export", "Wholesale" };

    /// <summary>
    /// <paramref name="count"/> rows, cycling through five regions and four categories so a grouped
    /// run has real groups to build without their number growing with the row count — a report with
    /// 100,000 groups is a different benchmark, and not a report.
    /// </summary>
    public static List<ReportRow> Rows(int count)
    {
        var rows = new List<ReportRow>(count);
        var start = new DateTime(2026, 1, 1, 0, 0, 0, DateTimeKind.Utc);

        for (var i = 0; i < count; i++)
        {
            rows.Add(new ReportRow
            {
                Region = Regions[i % Regions.Length],
                Category = Categories[i % Categories.Length],
                Product = "Product " + i,
                Reference = "REF-" + i.ToString("D7"),
                Quantity = 1 + (i % 97),
                UnitPrice = 5m + (i % 500),
                Discount = (i % 5) * 0.05m,
                SoldOn = start.AddDays(i % 365),
                IsExport = i % 7 == 0,
            });
        }

        return rows;
    }
}
