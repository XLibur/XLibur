using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel.ConditionalFormats;
using XLibur.Extensions;
using ColorType = DocumentFormat.OpenXml.Office2010.Excel.ColorType;
using ConditionalFormattingRule = DocumentFormat.OpenXml.Office2010.Excel.ConditionalFormattingRule;
using DataBar = DocumentFormat.OpenXml.Office2010.Excel.DataBar;
using Formula = DocumentFormat.OpenXml.Office.Excel.Formula;

namespace XLibur.Excel;

internal sealed class XLCFDataBarConverterExtension : IXLCFConverterExtension
{
    private static readonly Dictionary<ConditionalFormatValueObjectValues, ConditionalFormattingValueObjectTypeValues> CFValueToTypeMap =
        new Dictionary<ConditionalFormatValueObjectValues, ConditionalFormattingValueObjectTypeValues>
        {
            { ConditionalFormatValueObjectValues.Max, ConditionalFormattingValueObjectTypeValues.AutoMax },
            { ConditionalFormatValueObjectValues.Min, ConditionalFormattingValueObjectTypeValues.AutoMin },
            { ConditionalFormatValueObjectValues.Number, ConditionalFormattingValueObjectTypeValues.Numeric },
            { ConditionalFormatValueObjectValues.Percent, ConditionalFormattingValueObjectTypeValues.Percent },
            { ConditionalFormatValueObjectValues.Percentile, ConditionalFormattingValueObjectTypeValues.Percentile },
            { ConditionalFormatValueObjectValues.Formula, ConditionalFormattingValueObjectTypeValues.Formula },
        };

    public ConditionalFormattingRule Convert(IXLConditionalFormat cf, XLWorkbook.SaveContext context)
    {
        ConditionalFormattingRule conditionalFormattingRule = new ConditionalFormattingRule
        {
            Type = ConditionalFormatValues.DataBar,
            Id = ((XLConditionalFormat)cf).Id.WrapInBraces()
        };

        DataBar dataBar = new DataBar
        {
            MinLength = 0,
            MaxLength = 100,
            Gradient = cf.Gradient,
            ShowValue = !cf.ShowBarOnly,
        };

        if (cf.BarAxisPosition != XLDataBarAxisPosition.Automatic)
            dataBar.AxisPosition = cf.BarAxisPosition.ToOpenXml();

        var cfMinType = cf.ContentTypes.TryGetValue(1, out var contentType1)
            ? GetCFType(contentType1.ToOpenXml())
            : ConditionalFormattingValueObjectTypeValues.AutoMin;
        var cfMin = new ConditionalFormattingValueObject { Type = cfMinType };
        if (cf.Values.Any() && cf.Values[1]?.Value != null)
        {
            cfMin.Type = ConditionalFormattingValueObjectTypeValues.Numeric;
            cfMin.Append(new Formula { Text = cf.Values[1].Value });
        }

        var cfMaxType = cf.ContentTypes.TryGetValue(2, out var contentType2)
            ? GetCFType(contentType2.ToOpenXml())
            : ConditionalFormattingValueObjectTypeValues.AutoMax;
        var cfMax = new ConditionalFormattingValueObject { Type = cfMaxType };
        if (cf.Values.Count >= 2 && cf.Values[2]?.Value != null)
        {
            cfMax.Type = ConditionalFormattingValueObjectTypeValues.Numeric;
            cfMax.Append(new Formula { Text = cf.Values[2].Value });
        }

        var barAxisColor = new BarAxisColor();
        SetColorTypeProperties(barAxisColor, cf.BarAxisColor);

        var negativeColor = cf.Colors.Count == 2 ? cf.Colors[2] : cf.Colors[1];
        var negativeFillColor = new NegativeFillColor();
        SetColorTypeProperties(negativeFillColor, negativeColor);

        dataBar.Append(cfMin);
        dataBar.Append(cfMax);

        dataBar.Append(negativeFillColor);
        dataBar.Append(barAxisColor);

        conditionalFormattingRule.Append(dataBar);

        return conditionalFormattingRule;
    }

    private static ConditionalFormattingValueObjectTypeValues GetCFType(ConditionalFormatValueObjectValues value)
    {
        return CFValueToTypeMap[value];
    }

    private static void SetColorTypeProperties(ColorType target, XLColor color)
    {
        switch (color.ColorType)
        {
            case XLColorType.Color:
                target.Rgb = color.Color.ToHex();
                break;
            case XLColorType.Theme:
                target.Theme = System.Convert.ToUInt32(color.ThemeColor);
                if (color.ThemeTint != 0)
                    target.Tint = color.ThemeTint;
                break;
            case XLColorType.Indexed:
                target.Indexed = System.Convert.ToUInt32(color.Indexed);
                break;
        }
    }
}
