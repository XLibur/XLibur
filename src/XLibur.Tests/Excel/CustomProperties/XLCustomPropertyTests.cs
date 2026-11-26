using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using XLibur.Excel;
using System.Threading.Tasks;
using XLibur.Tests.Utils;

namespace XLibur.Tests.Excel.CustomProperties;

public class XLCustomPropertyTests
{
    [Test]
    public async Task NumericString_IsPreservedAsText()
    {
        using var wb = new XLWorkbook();
        wb.AddWorksheet("Sheet1");
        wb.CustomProperties.Add("OrderId", "12345");

        var prop = wb.CustomProperty("OrderId");

        await Assert.That(prop.Type).IsEqualTo(XLCustomPropertyType.Text);
        await Assert.That(prop.GetValue<string>()).IsEqualTo("12345");
    }

    [Test]
    public async Task NumericString_SurvivesRoundTrip()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            wb.AddWorksheet("Sheet1");
            wb.CustomProperties.Add("OrderId", "12345");
            wb.SaveAs(ms);
        }

        ms.Position = 0;
        using (var wb = new XLWorkbook(ms))
        {
            var prop = wb.CustomProperty("OrderId");
            await Assert.That(prop.Type).IsEqualTo(XLCustomPropertyType.Text);
            await Assert.That(prop.GetValue<string>()).IsEqualTo("12345");
        }
    }

    [Test]
    public async Task Double_IsStoredAsNumber()
    {
        using var wb = new XLWorkbook();
        wb.AddWorksheet("Sheet1");
        wb.CustomProperties.Add("Price", 99.99);

        var prop = wb.CustomProperty("Price");
        await Assert.That(prop.Type).IsEqualTo(XLCustomPropertyType.Number);
    }

    [Test]
    public async Task Integer_IsStoredAsNumber()
    {
        using var wb = new XLWorkbook();
        wb.AddWorksheet("Sheet1");
        wb.CustomProperties.Add("Count", 42);

        var prop = wb.CustomProperty("Count");
        await Assert.That(prop.Type).IsEqualTo(XLCustomPropertyType.Number);
    }

    [Test]
    public async Task Integer_property_from_a_file_survives_a_round_trip()
    {
        using var saved = TestHelper.LoadAndSave(@"TryToLoad\LO\xlsx\ceiling-floor.xlsx");

        await Assert.That(saved.ReadPart("docProps/custom.xml")).Contains("name=\"DocSecurity\"><vt:i4>0</vt:i4>");

        using var wb = new XLWorkbook(saved);
        var prop = wb.CustomProperty("DocSecurity");
        await Assert.That(prop.Type).IsEqualTo(XLCustomPropertyType.Number);
        await Assert.That(prop.Value).IsEqualTo(0);
    }

    [Test]
    public async Task Integer_added_in_code_is_written_as_an_integer()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            wb.AddWorksheet("Sheet1");
            wb.CustomProperties.Add("Count", 42);
            wb.SaveAs(ms);
        }

        await Assert.That(ms.ReadPart("docProps/custom.xml")).Contains("<vt:i4>42</vt:i4>");
    }

    [Test]
    [Arguments("i1", "-5", (sbyte)-5)]
    [Arguments("i2", "-300", (short)-300)]
    [Arguments("i4", "-70000", -70000)]
    [Arguments("i8", "9000000000", 9000000000L)]
    [Arguments("ui1", "200", (byte)200)]
    [Arguments("ui2", "60000", (ushort)60000)]
    [Arguments("ui4", "4000000000", 4000000000U)]
    [Arguments("ui8", "18000000000000000000", 18000000000000000000UL)]
    public async Task Integer_variant_types_load_as_numbers_and_keep_their_type(string vt, string text, object expected)
    {
        using var saved = RoundTripCustomProperties($"<property fmtid=\"{FormatId}\" pid=\"2\" name=\"P\"><vt:{vt}>{text}</vt:{vt}></property>");

        await Assert.That(saved.ReadPart("docProps/custom.xml")).Contains($"<vt:{vt}>{text}</vt:{vt}>");

        using var wb = new XLWorkbook(saved);
        var prop = wb.CustomProperty("P");
        await Assert.That(prop.Type).IsEqualTo(XLCustomPropertyType.Number);
        await Assert.That(prop.Value).IsEqualTo(expected);
    }

    [Test]
    [Arguments("<vt:lpstr>ansi text</vt:lpstr>")]
    [Arguments("<vt:decimal>12.5</vt:decimal>")]
    [Arguments("<vt:cy>1234.5600</vt:cy>")]
    [Arguments("<vt:r4>1.5</vt:r4>")]
    public async Task Unmodeled_variant_types_survive_a_round_trip(string variant)
    {
        using var saved = RoundTripCustomProperties($"<property fmtid=\"{FormatId}\" pid=\"2\" name=\"P\">{variant}</property>");

        await Assert.That(saved.ReadPart("docProps/custom.xml")).Contains($"name=\"P\">{variant}</");
    }

    [Test]
    public async Task Deleting_an_unmodeled_property_removes_it()
    {
        using var source = RoundTripCustomProperties(
            $"<property fmtid=\"{FormatId}\" pid=\"2\" name=\"P\"><vt:lpstr>ansi text</vt:lpstr></property>");
        using var wb = new XLWorkbook(source);
        wb.CustomProperties.Delete("P");

        using var saved = new MemoryStream();
        wb.SaveAs(saved);

        await Assert.That(saved.PartExists("docProps/custom.xml")).IsFalse();
    }

    [Test]
    public async Task A_property_added_in_code_replaces_an_unmodeled_one_of_the_same_name()
    {
        using var source = RoundTripCustomProperties(
            $"<property fmtid=\"{FormatId}\" pid=\"2\" name=\"P\"><vt:lpstr>ansi text</vt:lpstr></property>");
        using var wb = new XLWorkbook(source);
        wb.CustomProperties.Add("P", "replaced");

        using var saved = new MemoryStream();
        wb.SaveAs(saved);

        var xml = saved.ReadPart("docProps/custom.xml");
        await Assert.That(xml).Contains("<vt:lpwstr>replaced</vt:lpwstr>");
        await Assert.That(xml).DoesNotContain("ansi text");
    }

    [Test]
    public async Task Renaming_a_replacement_does_not_bring_back_the_unmodeled_property_it_replaced()
    {
        using var source = RoundTripCustomProperties(
            $"<property fmtid=\"{FormatId}\" pid=\"2\" name=\"P\"><vt:cy>1.5</vt:cy></property>");
        using var wb = new XLWorkbook(source);
        wb.CustomProperties.Add("P", "replacement");
        wb.CustomProperty("P").Name = "Q";

        using var saved = new MemoryStream();
        wb.SaveAs(saved);

        var xml = saved.ReadPart("docProps/custom.xml");
        await Assert.That(xml).Contains("name=\"Q\"><vt:lpwstr>replacement</vt:lpwstr>");
        await Assert.That(xml).DoesNotContain("name=\"P\"");
        await Assert.That(xml).DoesNotContain("vt:cy");
    }

    [Test]
    public async Task Renaming_onto_an_unmodeled_property_replaces_it_for_good()
    {
        using var source = RoundTripCustomProperties(
            $"<property fmtid=\"{FormatId}\" pid=\"2\" name=\"P\"><vt:cy>1.5</vt:cy></property>");
        using var wb = new XLWorkbook(source);
        wb.CustomProperties.Add("Q", "replacement");
        var property = wb.CustomProperty("Q");
        property.Name = "P";
        property.Name = "R";

        using var saved = new MemoryStream();
        wb.SaveAs(saved);

        var xml = saved.ReadPart("docProps/custom.xml");
        await Assert.That(xml).Contains("name=\"R\"><vt:lpwstr>replacement</vt:lpwstr>");
        await Assert.That(xml).DoesNotContain("vt:cy");
    }

    [Test]
    public async Task A_renamed_property_is_found_by_its_new_name_and_frees_its_old_one()
    {
        using var wb = new XLWorkbook();
        wb.CustomProperties.Add("P", "x");
        wb.CustomProperty("P").Name = "Q";
        wb.CustomProperties.Add("P", "y");

        await Assert.That(wb.CustomProperty("Q").Value).IsEqualTo("x");
        await Assert.That(wb.CustomProperty("P").Value).IsEqualTo("y");
    }

    [Test]
    public async Task A_renamed_property_is_saved_once_under_its_new_name_in_its_old_place()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            wb.AddWorksheet("Sheet1");
            wb.CustomProperties.Add("A", "a");
            wb.CustomProperties.Add("P", "x");
            wb.CustomProperties.Add("C", "c");
            wb.CustomProperty("P").Name = "Q";

            await Assert.That(string.Join(",", wb.CustomProperties.Select(p => p.Name))).IsEqualTo("A,Q,C");
            wb.SaveAs(ms);
        }

        var xml = ms.ReadPart("docProps/custom.xml");
        await Assert.That(xml).DoesNotContain("name=\"P\"");
        await Assert.That(xml.Split("name=\"Q\"").Length).IsEqualTo(2);
        await Assert.That(xml.IndexOf("name=\"A\"")).IsLessThan(xml.IndexOf("name=\"Q\""));
        await Assert.That(xml.IndexOf("name=\"Q\"")).IsLessThan(xml.IndexOf("name=\"C\""));
    }

    [Test]
    public async Task Every_property_can_be_renamed_while_enumerating_the_collection()
    {
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            wb.AddWorksheet("Sheet1");
            wb.CustomProperties.Add("A", "a");
            wb.CustomProperties.Add("B", "b");
            wb.CustomProperties.Add("C", "c");

            foreach (var p in wb.CustomProperties)
                p.Name += "_copy";

            await Assert.That(wb.CustomProperty("A_copy").Value).IsEqualTo("a");
            await Assert.That(wb.CustomProperty("B_copy").Value).IsEqualTo("b");
            await Assert.That(wb.CustomProperty("C_copy").Value).IsEqualTo("c");
            await Assert.That(() => wb.CustomProperty("A")).Throws<KeyNotFoundException>();
            await Assert.That(() => wb.CustomProperty("B")).Throws<KeyNotFoundException>();
            await Assert.That(() => wb.CustomProperty("C")).Throws<KeyNotFoundException>();
            wb.SaveAs(ms);
        }

        var xml = ms.ReadPart("docProps/custom.xml");
        await Assert.That(xml).DoesNotContain("name=\"A\"");
        await Assert.That(xml.IndexOf("name=\"A_copy\"")).IsGreaterThanOrEqualTo(0);
        await Assert.That(xml.IndexOf("name=\"A_copy\"")).IsLessThan(xml.IndexOf("name=\"B_copy\""));
        await Assert.That(xml.IndexOf("name=\"B_copy\"")).IsLessThan(xml.IndexOf("name=\"C_copy\""));
    }

    [Test]
    public async Task Deleting_a_renamed_property_by_its_new_name_removes_it()
    {
        using var wb = new XLWorkbook();
        wb.CustomProperties.Add("P", "x");
        wb.CustomProperty("P").Name = "Q";
        wb.CustomProperties.Delete("Q");

        await Assert.That(wb.CustomProperties).IsEmpty();
    }

    [Test]
    public async Task Renaming_onto_a_taken_name_throws_and_keeps_both_properties()
    {
        using var wb = new XLWorkbook();
        wb.CustomProperties.Add("P", "x");
        wb.CustomProperties.Add("Q", "y");

        await Assert.That(() => wb.CustomProperty("P").Name = "Q").Throws<System.ArgumentException>();
        wb.CustomProperty("P").Name = "P";

        await Assert.That(wb.CustomProperty("P").Value).IsEqualTo("x");
        await Assert.That(wb.CustomProperty("Q").Value).IsEqualTo("y");
    }

    [Test]
    [Arguments("<vt:lpwstr>cached</vt:lpwstr>")]
    [Arguments("<vt:cy>1.5</vt:cy>")]
    public async Task Link_target_survives_a_round_trip(string variant)
    {
        using var saved = RoundTripCustomProperties(
            $"<property fmtid=\"{FormatId}\" pid=\"2\" name=\"P\" linkTarget=\"MyName\">{variant}</property>");

        await Assert.That(saved.ReadPart("docProps/custom.xml")).Contains("linkTarget=\"MyName\"");
    }

    private const string FormatId = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}";

    private static MemoryStream RoundTripCustomProperties(string propertiesXml)
    {
        var source = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            wb.AddWorksheet("Sheet1");
            wb.CustomProperties.Add("Placeholder", "x");
            wb.SaveAs(source);
        }

        using (var document = SpreadsheetDocument.Open(source, true))
        {
            var xml = "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/custom-properties\" " +
                      "xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">" +
                      propertiesXml + "</Properties>";
            using var partStream = document.CustomFilePropertiesPart!.GetStream(FileMode.Create);
            using var writer = new StreamWriter(partStream);
            writer.Write(xml);
        }

        source.Position = 0;
        using var loaded = new XLWorkbook(source);
        var saved = new MemoryStream();
        loaded.SaveAs(saved);
        saved.Position = 0;
        return saved;
    }
}
