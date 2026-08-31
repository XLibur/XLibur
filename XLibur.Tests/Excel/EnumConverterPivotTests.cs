using System;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel;

namespace XLibur.Tests.Excel;

/// <summary>
/// Round-trip coverage for the nine pivot enums <see cref="EnumConverter"/> maps, folded in from
/// the pivot writer's own private string tables by spec 39. Every member, both directions, so a
/// value added to one side without the other fails here instead of drifting silently.
/// </summary>
public class EnumConverterPivotTests
{
    [Test]
    public async Task XLPivotSortType_round_trips()
    {
        foreach (var value in Enum.GetValues<XLPivotSortType>())
            await Assert.That(value.ToOpenXml().ToXLibur()).IsEqualTo(value);
    }

    [Test]
    public async Task XLPivotAxis_round_trips()
    {
        foreach (var value in Enum.GetValues<XLPivotAxis>())
            await Assert.That(value.ToOpenXml().ToXLibur()).IsEqualTo(value);
    }

    [Test]
    public async Task XLPivotItemType_round_trips()
    {
        foreach (var value in Enum.GetValues<XLPivotItemType>())
            await Assert.That(value.ToOpenXml().ToXLibur()).IsEqualTo(value);
    }

    [Test]
    public async Task XLPivotSummary_round_trips()
    {
        foreach (var value in Enum.GetValues<XLPivotSummary>())
            await Assert.That(value.ToOpenXml().ToXLibur()).IsEqualTo(value);
    }

    [Test]
    public async Task XLPivotCalculation_round_trips()
    {
        foreach (var value in Enum.GetValues<XLPivotCalculation>())
            await Assert.That(value.ToOpenXml().ToXLibur()).IsEqualTo(value);
    }

    [Test]
    public async Task XLPivotAreaType_round_trips()
    {
        foreach (var value in Enum.GetValues<XLPivotAreaType>())
            await Assert.That(value.ToOpenXml().ToXLibur()).IsEqualTo(value);
    }

    [Test]
    public async Task XLPivotFormatAction_round_trips()
    {
        foreach (var value in Enum.GetValues<XLPivotFormatAction>())
            await Assert.That(value.ToOpenXml().ToXLibur()).IsEqualTo(value);
    }

    [Test]
    public async Task XLPivotCfScope_round_trips()
    {
        foreach (var value in Enum.GetValues<XLPivotCfScope>())
            await Assert.That(value.ToOpenXml().ToXLibur()).IsEqualTo(value);
    }

    [Test]
    public async Task XLPivotCfRuleType_round_trips()
    {
        foreach (var value in Enum.GetValues<XLPivotCfRuleType>())
            await Assert.That(value.ToOpenXml().ToXLibur()).IsEqualTo(value);
    }

    /// <summary>
    /// <c>ST_DataConsolidateFunction</c> (a data field's <c>subtotal</c> attribute) spells the two
    /// population variants with a lower-case suffix; <c>ST_ItemType</c> (an item's <c>t</c>
    /// attribute) spells the same two concepts with an upper-case suffix. The pivot writer used to
    /// carry both spellings by hand in two separate private string tables — exactly the kind of
    /// one-character, easy-to-miss divergence spec 39 folds into <see cref="EnumConverter"/> so it
    /// can only be stated once per schema type.
    /// </summary>
    [Test]
    public async Task XLPivotSummary_population_variants_use_the_lower_case_spelling()
    {
        await Assert.That(new EnumValue<DataConsolidateFunctionValues>(XLPivotSummary.PopulationStandardDeviation.ToOpenXml()).InnerText)
            .IsEqualTo("stdDevp");
        await Assert.That(new EnumValue<DataConsolidateFunctionValues>(XLPivotSummary.PopulationVariance.ToOpenXml()).InnerText)
            .IsEqualTo("varp");
    }

    [Test]
    public async Task XLPivotItemType_population_variants_use_the_upper_case_spelling()
    {
        await Assert.That(new EnumValue<ItemValues>(XLPivotItemType.StdDevP.ToOpenXml()).InnerText)
            .IsEqualTo("stdDevP");
        await Assert.That(new EnumValue<ItemValues>(XLPivotItemType.VarP.ToOpenXml()).InnerText)
            .IsEqualTo("varP");
    }
}
