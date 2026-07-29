using System;

namespace XLibur.Report.Tests;

/// <summary>A line item, the shape most report fixtures bind.</summary>
public sealed class SaleItem
{
    public string Product { get; set; } = string.Empty;

    /// <summary>The outer key in grouping fixtures.</summary>
    public string Region { get; set; } = string.Empty;

    /// <summary>The inner key in nested-grouping fixtures.</summary>
    public string Category { get; set; } = string.Empty;

    public int Quantity { get; set; }

    public decimal UnitPrice { get; set; }

    public DateTime SoldOn { get; set; }

    public bool IsExport { get; set; }

    public string? Notes { get; set; }

    public decimal Total => Quantity * UnitPrice;
}

/// <summary>A container used to exercise <see cref="IXLTemplate.AddVariable(object)"/>.</summary>
public sealed class ReportHeader
{
    public string Company { get; set; } = string.Empty;

    public DateTime RunOn { get; set; }

    public int Year = 2026;
}
