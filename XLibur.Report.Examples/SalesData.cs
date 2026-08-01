using System;
using System.Collections.Generic;

namespace XLibur.Report.Examples;

/// <summary>One line of a sales report — the shape most of these examples bind.</summary>
/// <remarks>
/// An ordinary class with ordinary properties. A template refers to these by their C# names
/// (<c>{{ item.UnitPrice }}</c>), which is what the default engine's identity member renamer is for:
/// nothing has to be annotated or renamed to be bindable.
/// </remarks>
public sealed class Sale
{
    public string Region { get; init; } = string.Empty;

    public string Category { get; init; } = string.Empty;

    public string Product { get; init; } = string.Empty;

    public int Quantity { get; init; }

    public decimal UnitPrice { get; init; }

    public DateTime SoldOn { get; init; }

    public bool IsExport { get; init; }

    /// <summary>A computed property is bindable like any other.</summary>
    public decimal Total => Quantity * UnitPrice;
}

/// <summary>The data the examples run on.</summary>
public static class SalesData
{
    /// <summary>
    /// Twelve lines over three regions and two categories, deliberately out of order so that
    /// <c>&lt;&lt;Sort&gt;&gt;</c> and <c>&lt;&lt;Group&gt;&gt;</c> have something to do.
    /// </summary>
    public static List<Sale> Sales() => new()
    {
        Line("South", "Trade", "Rotary hoe", 4, 240.00m, 14, export: false),
        Line("North", "Retail", "Watering can", 32, 12.50m, 3, export: false),
        Line("East", "Trade", "Cultivator", 2, 480.00m, 21, export: true),
        Line("North", "Trade", "Seed drill", 3, 310.00m, 8, export: true),
        Line("South", "Retail", "Trowel", 96, 4.20m, 17, export: false),
        Line("East", "Retail", "Secateurs", 40, 18.75m, 22, export: false),
        Line("North", "Retail", "Dibber", 120, 2.80m, 5, export: false),
        Line("South", "Trade", "Bed frame", 11, 96.00m, 19, export: false),
        Line("East", "Trade", "Poly tunnel", 1, 1450.00m, 27, export: true),
        Line("North", "Trade", "Cloche", 18, 42.00m, 11, export: false),
        Line("South", "Retail", "Twine", 240, 1.15m, 24, export: false),
        Line("East", "Retail", "Kneeler", 26, 22.40m, 29, export: false),
    };

    private static Sale Line(
        string region, string category, string product, int quantity, decimal unitPrice, int day, bool export) =>
        new()
        {
            Region = region,
            Category = category,
            Product = product,
            Quantity = quantity,
            UnitPrice = unitPrice,
            SoldOn = new DateTime(2026, 3, day, 0, 0, 0, DateTimeKind.Utc),
            IsExport = export,
        };
}
