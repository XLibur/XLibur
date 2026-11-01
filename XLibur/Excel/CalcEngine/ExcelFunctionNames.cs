using System;
using System.Collections.Frozen;
using System.Linq;

namespace XLibur.Excel.CalcEngine;

/// <summary>
/// The names of Excel's own worksheet functions, whether XLibur implements them or not.
/// </summary>
/// <remarks>
/// <para>
/// A call by one of these names calls Excel's function even where a <c>LET</c> or <c>LAMBDA</c> binds
/// the name: Excel 365 calculates <c>LET(SUM,LAMBDA(x,x*2),SUM(3))</c> as 3, with <c>SUM</c> left as
/// it is. The function registry cannot say this, as it holds only the functions XLibur evaluates.
/// </para>
/// <para>
/// The functions newer than Excel 2007 are the keys of <see cref="XLConstants.FutureFunctionMap"/>,
/// the ones a file writes with <c>_xlfn.</c>. The rest are listed here: the table of Microsoft's "Excel functions
/// (alphabetical)" page, and the functions Excel 365's <c>WorksheetFunction</c> object exposes that
/// the table leaves out, such as the Thai text functions. Both were read on 2026-11-01.
/// </para>
/// </remarks>
internal static class ExcelFunctionNames
{
    private static readonly FrozenSet<string> Classic = new[]
    {
        "ABS", "ACCRINT", "ACCRINTM", "ACOS", "ACOSH", "ADDRESS", "AMORDEGRC", "AMORLINC", "AND", "AREAS",
        "ARRAYTOTEXT", "ASC", "ASIN", "ASINH", "ATAN", "ATAN2", "ATANH", "AVEDEV", "AVERAGE", "AVERAGEA",
        "AVERAGEIF", "AVERAGEIFS", "BAHTTEXT", "BESSELI", "BESSELJ", "BESSELK", "BESSELY", "BETADIST",
        "BETAINV", "BIN2DEC", "BIN2HEX", "BIN2OCT", "BINOMDIST", "CALL", "CEILING", "CELL", "CHAR",
        "CHIDIST", "CHIINV", "CHITEST", "CHOOSE", "CLEAN", "CODE", "COLUMN", "COLUMNS", "COMBIN",
        "COMPLEX", "CONCATENATE", "CONFIDENCE", "CONVERT", "CORREL", "COS", "COSH", "COUNT", "COUNTA",
        "COUNTBLANK", "COUNTIF", "COUNTIFS", "COUPDAYBS", "COUPDAYS", "COUPDAYSNC", "COUPNCD", "COUPNUM",
        "COUPPCD", "COVAR", "CRITBINOM", "CUBEKPIMEMBER", "CUBEMEMBER", "CUBEMEMBERPROPERTY",
        "CUBERANKEDMEMBER", "CUBESET", "CUBESETCOUNT", "CUBEVALUE", "CUMIPMT", "CUMPRINC", "DATE",
        "DATEDIF", "DATEVALUE", "DAVERAGE", "DAY", "DAYS360", "DB", "DBCS", "DCOUNT", "DCOUNTA", "DDB",
        "DEC2BIN", "DEC2HEX", "DEC2OCT", "DEGREES", "DELTA", "DEVSQ", "DGET", "DISC", "DMAX", "DMIN",
        "DOLLAR", "DOLLARDE", "DOLLARFR", "DPRODUCT", "DSTDEV", "DSTDEVP", "DSUM", "DURATION", "DVAR",
        "DVARP", "EDATE", "EFFECT", "ENCODEURL", "EOMONTH", "ERF", "ERFC", "ERROR.TYPE", "EUROCONVERT",
        "EVEN", "EXACT", "EXP", "EXPONDIST", "FACT", "FACTDOUBLE", "FALSE", "FDIST", "FIND", "FINDB",
        "FINV", "FISHER", "FISHERINV", "FIXED", "FLOOR", "FORECAST", "FREQUENCY", "FTEST", "FV",
        "FVSCHEDULE", "GAMMADIST", "GAMMAINV", "GAMMALN", "GCD", "GEOMEAN", "GESTEP", "GETPIVOTDATA",
        "GROWTH", "HARMEAN", "HEX2BIN", "HEX2DEC", "HEX2OCT", "HLOOKUP", "HOUR", "HYPERLINK",
        "HYPGEOMDIST", "IF", "IFERROR", "IMABS", "IMAGINARY", "IMARGUMENT", "IMCONJUGATE", "IMCOS",
        "IMDIV", "IMEXP", "IMLN", "IMLOG10", "IMLOG2", "IMPOWER", "IMPRODUCT", "IMREAL", "IMSIN", "IMSQRT",
        "IMSUB", "IMSUM", "INDEX", "INDIRECT", "INFO", "INT", "INTERCEPT", "INTRATE", "IPMT", "IRR",
        "ISBLANK", "ISERR", "ISERROR", "ISEVEN", "ISLOGICAL", "ISNA", "ISNONTEXT", "ISNUMBER",
        "ISO.CEILING", "ISODD", "ISPMT", "ISREF", "ISTEXT", "ISTHAIDIGIT", "JIS", "KURT", "LARGE", "LCM",
        "LEFT", "LEFTB", "LEN", "LENB", "LINEST", "LN", "LOG", "LOG10", "LOGEST", "LOGINV", "LOGNORMDIST",
        "LOOKUP", "LOWER", "MATCH", "MAX", "MAXA", "MDETERM", "MDURATION", "MEDIAN", "MID", "MIDB", "MIN",
        "MINA", "MINUTE", "MINVERSE", "MIRR", "MMULT", "MOD", "MODE", "MONTH", "MROUND", "MULTINOMIAL",
        "N", "NA", "NEGBINOMDIST", "NETWORKDAYS", "NETWORKDAYS.INTL", "NOMINAL", "NORMDIST", "NORMINV",
        "NORMSDIST", "NORMSINV", "NOT", "NOW", "NPER", "NPV", "OCT2BIN", "OCT2DEC", "OCT2HEX", "ODD",
        "ODDFPRICE", "ODDFYIELD", "ODDLPRICE", "ODDLYIELD", "OFFSET", "OR", "PEARSON", "PERCENTILE",
        "PERCENTRANK", "PERMUT", "PHONETIC", "PI", "PMT", "POISSON", "POWER", "PPMT", "PRICE", "PRICEDISC",
        "PRICEMAT", "PROB", "PRODUCT", "PROPER", "PV", "QUARTILE", "QUOTIENT", "RADIANS", "RAND",
        "RANDBETWEEN", "RANK", "RATE", "RECEIVED", "REGISTER.ID", "REPLACE", "REPLACEB", "REPT", "RIGHT",
        "RIGHTB", "ROMAN", "ROUND", "ROUNDBAHTDOWN", "ROUNDBAHTUP", "ROUNDDOWN", "ROUNDUP", "ROW", "ROWS",
        "RSQ", "RTD", "SEARCH", "SEARCHB", "SECOND", "SERIESSUM", "SIGN", "SIN", "SINGLE", "SINH", "SKEW",
        "SLN", "SLOPE", "SMALL", "SQRT", "SQRTPI", "STANDARDIZE", "STDEV", "STDEVA", "STDEVP", "STDEVPA",
        "STEYX", "STOCKHISTORY", "SUBSTITUTE", "SUBTOTAL", "SUM", "SUMIF", "SUMIFS", "SUMPRODUCT", "SUMSQ",
        "SUMX2MY2", "SUMX2PY2", "SUMXMY2", "SYD", "T", "TAN", "TANH", "TBILLEQ", "TBILLPRICE",
        "TBILLYIELD", "TDIST", "TEXT", "THAIDAYOFWEEK", "THAIDIGIT", "THAIMONTHOFYEAR", "THAINUMSOUND",
        "THAINUMSTRING", "THAISTRINGLENGTH", "THAIYEAR", "TIME", "TIMEVALUE", "TINV", "TODAY", "TRANSPOSE",
        "TREND", "TRIM", "TRIMMEAN", "TRUE", "TRUNC", "TTEST", "TYPE", "UPPER", "USDOLLAR", "VALUE",
        "VALUETOTEXT", "VAR", "VARA", "VARP", "VARPA", "VDB", "VLOOKUP", "WEEKDAY", "WEEKNUM", "WEIBULL",
        "WORKDAY", "WORKDAY.INTL", "XIRR", "XMATCH", "XNPV", "YEAR", "YEARFRAC", "YIELD", "YIELDDISC",
        "YIELDMAT", "ZTEST",
    }.ToFrozenSet(StringComparer.OrdinalIgnoreCase);

    /// <summary>Every name <see cref="IsBuiltIn(string)"/> is true for, in one set.</summary>
    private static readonly Lazy<FrozenSet<string>> All = new(() =>
        Classic.Concat(XLConstants.FutureFunctionMap.Value.Keys).ToFrozenSet(StringComparer.OrdinalIgnoreCase));

    /// <summary>Whether <paramref name="name"/>, without a future-function prefix, names one of Excel's functions.</summary>
    internal static bool IsBuiltIn(string name) =>
        Classic.Contains(name) || XLConstants.FutureFunctionMap.Value.ContainsKey(name);

    /// <inheritdoc cref="IsBuiltIn(string)"/>
    /// <remarks>Looked up without making a string of the name where the framework can.</remarks>
    internal static bool IsBuiltIn(ReadOnlySpan<char> name)
    {
#if NET9_0_OR_GREATER
        return All.Value.GetAlternateLookup<ReadOnlySpan<char>>().Contains(name);
#else
        return All.Value.Contains(name.ToString());
#endif
    }
}
