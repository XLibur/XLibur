namespace XLibur.Excel;

/// <summary>
/// What a <see cref="IXLPivotCache"/> reads its records from.
/// </summary>
/// <remarks>
/// Only <see cref="Range"/> and <see cref="Name"/> can resolve to a worksheet in this workbook.
/// XLibur cannot read the others, so a cache with one of those sources keeps whatever records the
/// file was saved with — <see cref="IXLPivotCache.Refresh"/> throws for them.
/// </remarks>
public enum XLPivotSourceKind
{
    /// <summary>A direct cell area on a sheet in this workbook.</summary>
    Range,

    /// <summary>A table or a book-scoped defined name in this workbook.</summary>
    Name,

    /// <summary>Several ranges consolidated into one source.</summary>
    Consolidation,

    /// <summary>The workbook's scenario data.</summary>
    Scenario,

    /// <summary>A range in a different workbook.</summary>
    ExternalWorkbook,

    /// <summary>An external data connection — a database, a query, a cube.</summary>
    Connection,
}
