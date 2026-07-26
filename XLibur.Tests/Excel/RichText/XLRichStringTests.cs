using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using XLibur.Excel;
using XLibur.Excel.IO;
using System.Threading.Tasks;

namespace XLibur.Tests.Excel.RichText;
/// <summary>
///     This is a test class for XLRichStringTests and is intended
///     to contain all XLRichStringTests Unit Tests
/// </summary>
public class XLRichStringTests
{
    [Test]
    public async Task AccessRichTextTest1()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var cell = ws.Cell(1, 1);
        cell.CreateRichText().AddText("12");

        var richText = cell.GetRichText();

        await Assert.That(richText.ToString()).IsEqualTo("12");

        richText.AddText("34");

        await Assert.That(cell.GetText()).IsEqualTo("1234");
    }

    /// <summary>
    ///     A test for AddText
    /// </summary>
    [Test]
    public async Task AddTextTest1()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var cell = ws.Cell(1, 1);
        var richString = cell.CreateRichText();

        const string text = "Hello";
        richString.AddText(text).SetBold().SetFontColor(XLColor.Red);

        await Assert.That(text).IsEqualTo(cell.GetText());
        await Assert.That(true).IsEqualTo(cell.GetRichText().First().Bold);
        await Assert.That(XLColor.Red).IsEqualTo(cell.GetRichText().First().FontColor);

        await Assert.That(richString.Count).IsEqualTo(1);

        richString.AddText("World");
        await Assert.That(text).IsEqualTo(richString.First().Text).Because("Item in collection is not the same as the one returned");
    }

    [Test]
    public async Task AddTextTest2()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var cell = ws.Cell(1, 1);
        const int number = 123;

        cell.SetValue(number).Style
            .Font.SetBold()
            .Font.SetFontColor(XLColor.Red);

        var text = number.ToString();

        await Assert.That(text).IsEqualTo(cell.GetRichText().ToString());
        await Assert.That(true).IsEqualTo(cell.GetRichText().First().Bold);
        await Assert.That(XLColor.Red).IsEqualTo(cell.GetRichText().First().FontColor);

        await Assert.That(cell.GetRichText().Count).IsEqualTo(1);

        cell.GetRichText().AddText("World");
        await Assert.That(text).IsEqualTo(cell.GetRichText().First().Text).Because("Item in collection is not the same as the one returned");
    }

    [Test]
    public async Task AddTextTest3()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var cell = ws.Cell(1, 1);
        const int number = 123;
        cell.Value = number;
        cell.Style
            .Font.SetBold()
            .Font.SetFontColor(XLColor.Red);

        var text = number.ToString();

        await Assert.That(text).IsEqualTo(cell.GetRichText().ToString());
        await Assert.That(true).IsEqualTo(cell.GetRichText().First().Bold);
        await Assert.That(XLColor.Red).IsEqualTo(cell.GetRichText().First().FontColor);

        await Assert.That(cell.GetRichText().Count).IsEqualTo(1);

        cell.GetRichText().AddText("World");
        await Assert.That(text).IsEqualTo(cell.GetRichText().First().Text).Because("Item in collection is not the same as the one returned");
    }

    /// <summary>
    /// A test for Clear
    /// </summary>
    [Test]
    public async Task ClearTest()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var richString = ws.Cell(1, 1).GetRichText();

        richString.AddText("Hello");
        richString.AddText(" ");
        richString.AddText("World!");

        richString.ClearText();
        var expected = String.Empty;
        var actual = richString.ToString();
        await Assert.That(actual).IsEqualTo(expected);

        await Assert.That(richString.Count).IsEqualTo(0);
    }

    [Test]
    public async Task CountTest()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var richString = ws.Cell(1, 1).GetRichText();

        richString.AddText("Hello");
        richString.AddText(" ");
        richString.AddText("World!");

        await Assert.That(richString.Count).IsEqualTo(3);
    }

    [Test]
    public async Task HasRichTextTest1()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var cell = ws.Cell(1, 1);
        cell.GetRichText().AddText("123");

        await Assert.That(cell.HasRichText).IsEqualTo(true);

        cell.Value = "123";

        await Assert.That(cell.HasRichText).IsEqualTo(false);

        cell.GetRichText().AddText("123");

        await Assert.That(cell.HasRichText).IsEqualTo(true);

        cell.Value = 123;

        await Assert.That(cell.HasRichText).IsEqualTo(false);

        cell.GetRichText().AddText("123");

        await Assert.That(cell.HasRichText).IsEqualTo(true);

        cell.SetValue("123");

        await Assert.That(cell.HasRichText).IsEqualTo(false);
    }

    /// <summary>
    ///     A test for Characters
    /// </summary>
    [Test]
    public async Task Substring_All_From_OneString()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var richString = ws.Cell(1, 1).GetRichText();

        richString.AddText("Hello");

        var actual = richString.Substring(0);

        await Assert.That(actual.First()).IsEqualTo(richString.First());

        await Assert.That(actual.Count).IsEqualTo(1);

        actual.First().SetBold();

        await Assert.That(ws.Cell(1, 1).GetRichText().First().Bold).IsEqualTo(true);
    }

    [Test]
    public async Task Substring_All_From_ThreeStrings()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var richString = ws.Cell(1, 1).GetRichText();

        richString.AddText("Good Morning");
        richString.AddText(" my ");
        richString.AddText("neighbors!");

        var actual = richString.Substring(0);

        await Assert.That(actual.ElementAt(0)).IsEqualTo(richString.ElementAt(0));
        await Assert.That(actual.ElementAt(1)).IsEqualTo(richString.ElementAt(1));
        await Assert.That(actual.ElementAt(2)).IsEqualTo(richString.ElementAt(2));

        await Assert.That(actual.Count).IsEqualTo(3);
        await Assert.That(richString.Count).IsEqualTo(3);

        actual.First().SetBold();

        await Assert.That(ws.Cell(1, 1).GetRichText().First().Bold).IsEqualTo(true);
        await Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(1).Bold).IsEqualTo(false);
        await Assert.That(ws.Cell(1, 1).GetRichText().Last().Bold).IsEqualTo(false);
    }

    [Test]
    public async Task Substring_From_OneString_End()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var richString = ws.Cell(1, 1).GetRichText();

        richString.AddText("Hello");

        var actual = richString.Substring(2);

        await Assert.That(actual.Count).IsEqualTo(1); // substring was in one piece

        await Assert.That(richString.Count).IsEqualTo(2); // The text was split because of the substring

        await Assert.That(actual.First().Text).IsEqualTo("llo");

        await Assert.That(richString.First().Text).IsEqualTo("He");
        await Assert.That(richString.Last().Text).IsEqualTo("llo");

        actual.First().SetBold();

        await Assert.That(ws.Cell(1, 1).GetRichText().First().Bold).IsEqualTo(false);
        await Assert.That(ws.Cell(1, 1).GetRichText().Last().Bold).IsEqualTo(true);

        richString.Last().SetItalic();

        await Assert.That(ws.Cell(1, 1).GetRichText().First().Italic).IsEqualTo(false);
        await Assert.That(ws.Cell(1, 1).GetRichText().Last().Italic).IsEqualTo(true);

        await Assert.That(actual.First().Italic).IsEqualTo(true);

        richString.SetFontSize(20);

        await Assert.That(ws.Cell(1, 1).GetRichText().First().FontSize).IsEqualTo(20);
        await Assert.That(ws.Cell(1, 1).GetRichText().Last().FontSize).IsEqualTo(20);

        await Assert.That(actual.First().FontSize).IsEqualTo(20);
    }

    [Test]
    public async Task Substring_From_OneString_Middle()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var richString = ws.Cell(1, 1).GetRichText();

        richString.AddText("Hello");

        var actual = richString.Substring(2, 2);

        await Assert.That(actual.Count).IsEqualTo(1); // substring was in one piece

        await Assert.That(richString.Count).IsEqualTo(3); // The text was split because of the substring

        await Assert.That(actual.First().Text).IsEqualTo("ll");

        await Assert.That(richString.First().Text).IsEqualTo("He");
        await Assert.That(richString.ElementAt(1).Text).IsEqualTo("ll");
        await Assert.That(richString.Last().Text).IsEqualTo("o");

        actual.First().SetBold();

        await Assert.That(ws.Cell(1, 1).GetRichText().First().Bold).IsEqualTo(false);
        await Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(1).Bold).IsEqualTo(true);
        await Assert.That(ws.Cell(1, 1).GetRichText().Last().Bold).IsEqualTo(false);

        richString.Last().SetItalic();

        await Assert.That(ws.Cell(1, 1).GetRichText().First().Italic).IsEqualTo(false);
        await Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(1).Italic).IsEqualTo(false);
        await Assert.That(ws.Cell(1, 1).GetRichText().Last().Italic).IsEqualTo(true);

        await Assert.That(actual.First().Italic).IsEqualTo(false);

        richString.SetFontSize(20);

        await Assert.That(ws.Cell(1, 1).GetRichText().First().FontSize).IsEqualTo(20);
        await Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(1).FontSize).IsEqualTo(20);
        await Assert.That(ws.Cell(1, 1).GetRichText().Last().FontSize).IsEqualTo(20);

        await Assert.That(actual.First().FontSize).IsEqualTo(20);
    }

    [Test]
    public async Task Substring_From_OneString_Start()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var richString = ws.Cell(1, 1).GetRichText();

        richString.AddText("Hello");

        var actual = richString.Substring(0, 2);

        await Assert.That(actual.Count).IsEqualTo(1); // substring was in one piece

        await Assert.That(richString.Count).IsEqualTo(2); // The text was split because of the substring

        await Assert.That(actual.First().Text).IsEqualTo("He");

        await Assert.That(richString.First().Text).IsEqualTo("He");
        await Assert.That(richString.Last().Text).IsEqualTo("llo");

        actual.First().SetBold();

        await Assert.That(ws.Cell(1, 1).GetRichText().First().Bold).IsEqualTo(true);
        await Assert.That(ws.Cell(1, 1).GetRichText().Last().Bold).IsEqualTo(false);

        richString.Last().SetItalic();

        await Assert.That(ws.Cell(1, 1).GetRichText().First().Italic).IsEqualTo(false);
        await Assert.That(ws.Cell(1, 1).GetRichText().Last().Italic).IsEqualTo(true);

        await Assert.That(actual.First().Italic).IsEqualTo(false);

        richString.SetFontSize(20);

        await Assert.That(ws.Cell(1, 1).GetRichText().First().FontSize).IsEqualTo(20);
        await Assert.That(ws.Cell(1, 1).GetRichText().Last().FontSize).IsEqualTo(20);

        await Assert.That(actual.First().FontSize).IsEqualTo(20);
    }

    [Test]
    public async Task Substring_From_ThreeStrings_End1()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var richString = ws.Cell(1, 1).GetRichText();

        richString.AddText("Good Morning");
        richString.AddText(" my ");
        richString.AddText("neighbors!");

        var actual = richString.Substring(21);

        await Assert.That(actual.Count).IsEqualTo(1); // substring was in one piece

        await Assert.That(richString.Count).IsEqualTo(4); // The text was split because of the substring

        await Assert.That(actual.First().Text).IsEqualTo("bors!");

        await Assert.That(richString.ElementAt(0).Text).IsEqualTo("Good Morning");
        await Assert.That(richString.ElementAt(1).Text).IsEqualTo(" my ");
        await Assert.That(richString.ElementAt(2).Text).IsEqualTo("neigh");
        await Assert.That(richString.ElementAt(3).Text).IsEqualTo("bors!");

        actual.First().SetBold();

        await Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(0).Bold).IsEqualTo(false);
        await Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(1).Bold).IsEqualTo(false);
        await Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(2).Bold).IsEqualTo(false);
        await Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(3).Bold).IsEqualTo(true);

        richString.Last().SetItalic();

        await Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(0).Italic).IsEqualTo(false);
        await Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(1).Italic).IsEqualTo(false);
        await Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(2).Italic).IsEqualTo(false);
        await Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(3).Italic).IsEqualTo(true);

        await Assert.That(actual.First().Italic).IsEqualTo(true);

        richString.SetFontSize(20);

        await Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(0).FontSize).IsEqualTo(20);
        await Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(1).FontSize).IsEqualTo(20);
        await Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(2).FontSize).IsEqualTo(20);
        await Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(3).FontSize).IsEqualTo(20);

        await Assert.That(actual.First().FontSize).IsEqualTo(20);
    }

    [Test]
    public async Task Substring_From_ThreeStrings_End2()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var richString = ws.Cell(1, 1).GetRichText();

        richString.AddText("Good Morning");
        richString.AddText(" my ");
        richString.AddText("neighbors!");

        var actual = richString.Substring(13);

        await Assert.That(actual.Count).IsEqualTo(2);

        await Assert.That(richString.Count).IsEqualTo(4); // The text was split because of the substring

        await Assert.That(actual.ElementAt(0).Text).IsEqualTo("my ");
        await Assert.That(actual.ElementAt(1).Text).IsEqualTo("neighbors!");

        await Assert.That(richString.ElementAt(0).Text).IsEqualTo("Good Morning");
        await Assert.That(richString.ElementAt(1).Text).IsEqualTo(" ");
        await Assert.That(richString.ElementAt(2).Text).IsEqualTo("my ");
        await Assert.That(richString.ElementAt(3).Text).IsEqualTo("neighbors!");

        actual.ElementAt(1).SetBold();

        await Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(0).Bold).IsEqualTo(false);
        await Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(1).Bold).IsEqualTo(false);
        await Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(2).Bold).IsEqualTo(false);
        await Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(3).Bold).IsEqualTo(true);

        richString.Last().SetItalic();

        await Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(0).Italic).IsEqualTo(false);
        await Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(1).Italic).IsEqualTo(false);
        await Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(2).Italic).IsEqualTo(false);
        await Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(3).Italic).IsEqualTo(true);

        await Assert.That(actual.ElementAt(0).Italic).IsEqualTo(false);
        await Assert.That(actual.ElementAt(1).Italic).IsEqualTo(true);

        richString.SetFontSize(20);

        await Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(0).FontSize).IsEqualTo(20);
        await Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(1).FontSize).IsEqualTo(20);
        await Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(2).FontSize).IsEqualTo(20);
        await Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(3).FontSize).IsEqualTo(20);

        await Assert.That(actual.ElementAt(0).FontSize).IsEqualTo(20);
        await Assert.That(actual.ElementAt(1).FontSize).IsEqualTo(20);
    }

    [Test]
    public async Task Substring_From_ThreeStrings_Mid1()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var richString = ws.Cell(1, 1).GetRichText();

        richString.AddText("Good Morning");
        richString.AddText(" my ");
        richString.AddText("neighbors!");

        var actual = richString.Substring(5, 10);

        await Assert.That(actual.Count).IsEqualTo(2);

        await Assert.That(richString.Count).IsEqualTo(5); // The text was split because of the substring

        await Assert.That(actual.ElementAt(0).Text).IsEqualTo("Morning");
        await Assert.That(actual.ElementAt(1).Text).IsEqualTo(" my");

        await Assert.That(richString.ElementAt(0).Text).IsEqualTo("Good ");
        await Assert.That(richString.ElementAt(1).Text).IsEqualTo("Morning");
        await Assert.That(richString.ElementAt(2).Text).IsEqualTo(" my");
        await Assert.That(richString.ElementAt(3).Text).IsEqualTo(" ");
        await Assert.That(richString.ElementAt(4).Text).IsEqualTo("neighbors!");
    }

    [Test]
    public async Task Substring_From_ThreeStrings_Mid2()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var richString = ws.Cell(1, 1).GetRichText();

        richString.AddText("Good Morning");
        richString.AddText(" my ");
        richString.AddText("neighbors!");

        var actual = richString.Substring(5, 15);

        await Assert.That(actual.Count).IsEqualTo(3);

        await Assert.That(richString.Count).IsEqualTo(5); // The text was split because of the substring

        await Assert.That(actual.ElementAt(0).Text).IsEqualTo("Morning");
        await Assert.That(actual.ElementAt(1).Text).IsEqualTo(" my ");
        await Assert.That(actual.ElementAt(2).Text).IsEqualTo("neig");

        await Assert.That(richString.ElementAt(0).Text).IsEqualTo("Good ");
        await Assert.That(richString.ElementAt(1).Text).IsEqualTo("Morning");
        await Assert.That(richString.ElementAt(2).Text).IsEqualTo(" my ");
        await Assert.That(richString.ElementAt(3).Text).IsEqualTo("neig");
        await Assert.That(richString.ElementAt(4).Text).IsEqualTo("hbors!");
    }

    [Test]
    public async Task Substring_From_ThreeStrings_Start1()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var richString = ws.Cell(1, 1).GetRichText();

        richString.AddText("Good Morning");
        richString.AddText(" my ");
        richString.AddText("neighbors!");

        var actual = richString.Substring(0, 4);

        await Assert.That(actual.Count).IsEqualTo(1); // substring was in one piece

        await Assert.That(richString.Count).IsEqualTo(4); // The text was split because of the substring

        await Assert.That(actual.First().Text).IsEqualTo("Good");

        await Assert.That(richString.ElementAt(0).Text).IsEqualTo("Good");
        await Assert.That(richString.ElementAt(1).Text).IsEqualTo(" Morning");
        await Assert.That(richString.ElementAt(2).Text).IsEqualTo(" my ");
        await Assert.That(richString.ElementAt(3).Text).IsEqualTo("neighbors!");

        actual.First().SetBold();

        await Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(0).Bold).IsEqualTo(true);
        await Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(1).Bold).IsEqualTo(false);
        await Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(2).Bold).IsEqualTo(false);
        await Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(3).Bold).IsEqualTo(false);

        richString.First().SetItalic();

        await Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(0).Italic).IsEqualTo(true);
        await Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(1).Italic).IsEqualTo(false);
        await Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(2).Italic).IsEqualTo(false);
        await Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(3).Italic).IsEqualTo(false);

        await Assert.That(actual.First().Italic).IsEqualTo(true);

        richString.SetFontSize(20);

        await Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(0).FontSize).IsEqualTo(20);
        await Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(1).FontSize).IsEqualTo(20);
        await Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(2).FontSize).IsEqualTo(20);
        await Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(3).FontSize).IsEqualTo(20);

        await Assert.That(actual.First().FontSize).IsEqualTo(20);
    }

    [Test]
    public async Task Substring_From_ThreeStrings_Start2()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var richString = ws.Cell(1, 1).GetRichText();

        richString.AddText("Good Morning");
        richString.AddText(" my ");
        richString.AddText("neighbors!");

        var actual = richString.Substring(0, 15);

        await Assert.That(actual.Count).IsEqualTo(2);

        await Assert.That(richString.Count).IsEqualTo(4); // The text was split because of the substring

        await Assert.That(actual.ElementAt(0).Text).IsEqualTo("Good Morning");
        await Assert.That(actual.ElementAt(1).Text).IsEqualTo(" my");

        await Assert.That(richString.ElementAt(0).Text).IsEqualTo("Good Morning");
        await Assert.That(richString.ElementAt(1).Text).IsEqualTo(" my");
        await Assert.That(richString.ElementAt(2).Text).IsEqualTo(" ");
        await Assert.That(richString.ElementAt(3).Text).IsEqualTo("neighbors!");

        actual.ElementAt(1).SetBold();

        await Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(0).Bold).IsEqualTo(false);
        await Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(1).Bold).IsEqualTo(true);
        await Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(2).Bold).IsEqualTo(false);
        await Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(3).Bold).IsEqualTo(false);

        richString.First().SetItalic();

        await Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(0).Italic).IsEqualTo(true);
        await Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(1).Italic).IsEqualTo(false);
        await Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(2).Italic).IsEqualTo(false);
        await Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(3).Italic).IsEqualTo(false);

        await Assert.That(actual.ElementAt(0).Italic).IsEqualTo(true);
        await Assert.That(actual.ElementAt(1).Italic).IsEqualTo(false);

        richString.SetFontSize(20);

        await Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(0).FontSize).IsEqualTo(20);
        await Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(1).FontSize).IsEqualTo(20);
        await Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(2).FontSize).IsEqualTo(20);
        await Assert.That(ws.Cell(1, 1).GetRichText().ElementAt(3).FontSize).IsEqualTo(20);

        await Assert.That(actual.ElementAt(0).FontSize).IsEqualTo(20);
        await Assert.That(actual.ElementAt(1).FontSize).IsEqualTo(20);
    }

    [Test]
    public async Task Substring_IndexOutsideRange1()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var richString = ws.Cell(1, 1).GetRichText();

        richString.AddText("Hello");

        await Assert.That(() => richString.Substring(50)).Throws<ArgumentOutOfRangeException>();
    }

    [Test]
    public async Task Substring_IndexOutsideRange2()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var richString = ws.Cell(1, 1).GetRichText();

        richString.AddText("Hello");
        richString.AddText("World");

        await Assert.That(() => richString.Substring(50)).Throws<ArgumentOutOfRangeException>();
    }

    [Test]
    public async Task Substring_IndexOutsideRange3()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var richString = ws.Cell(1, 1).GetRichText();

        richString.AddText("Hello");

        await Assert.That(() => richString.Substring(1, 10)).Throws<ArgumentOutOfRangeException>();
    }

    [Test]
    public async Task Substring_IndexOutsideRange4()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var richString = ws.Cell(1, 1).GetRichText();

        richString.AddText("Hello");
        richString.AddText("World");

        await Assert.That(() => richString.Substring(5, 20)).Throws<ArgumentOutOfRangeException>();
    }

    [Test]
    public async Task CopyFrom_DoesCopy()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var original = ws.Cell(1, 1).GetRichText();
        original
            .AddText("Hello").SetFontSize(15).SetFontColor(XLColor.Red)
            .AddText("World").SetFontSize(7).SetFontColor(XLColor.Blue);

        var otherCell = ws.Cell(1, 2);
        var otherRichText = otherCell.GetRichText();
        otherRichText.CopyFrom(original);

        await Assert.That(otherCell.Value).IsEqualTo("HelloWorld");
        await Assert.That(otherRichText.Count).IsEqualTo(2);
        await Assert.That(otherRichText.First().FontColor).IsEqualTo(XLColor.Red);
        await Assert.That(otherRichText.Last().FontColor).IsEqualTo(XLColor.Blue);
    }

    /// <summary>
    ///     A test for ToString
    /// </summary>
    [Test]
    public async Task ToStringTest()
    {
        var ws = new XLWorkbook().Worksheets.Add("Sheet1");
        var richString = ws.Cell(1, 1).GetRichText();

        richString.AddText("Hello");
        richString.AddText(" ");
        richString.AddText("World");
        var expected = "Hello World";
        var actual = richString.ToString();
        await Assert.That(actual).IsEqualTo(expected);

        richString.AddText("!");
        expected = "Hello World!";
        actual = richString.ToString();
        await Assert.That(actual).IsEqualTo(expected);

        richString.ClearText();
        expected = String.Empty;
        actual = richString.ToString();
        await Assert.That(actual).IsEqualTo(expected);
    }

    [Test]
    [Property("Description", "See #1361")]
    public async Task CanClearInlinedRichText()
    {
        using var outputStream = new MemoryStream();
        using (var inputStream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\InlinedRichText\ChangeRichText\inputfile.xlsx")))
        using (var workbook = new XLWorkbook(inputStream))
        {
            workbook.Worksheets.First().Cell("A1").Value = "";
            workbook.SaveAs(outputStream);
        }

        using (var wb = new XLWorkbook(outputStream))
        {
            await Assert.That(wb.Worksheets.First().Cell("A1").Value).IsEqualTo("");
        }
    }

    [Test]
    public async Task CanChangeInlinedRichText()
    {
        static async Task AssertRichText(IXLRichText richText)
        {
            await Assert.That(richText).IsNotNull();
            await Assert.That(richText.Any()).IsTrue();
            await Assert.That(richText.ElementAt(2).Text).IsEqualTo("3");
            await Assert.That(richText.ElementAt(2).FontColor).IsEqualTo(XLColor.Red);
        }

        using var outputStream = new MemoryStream();
        using (var inputStream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\InlinedRichText\ChangeRichText\inputfile.xlsx")))
        using (var workbook = new XLWorkbook(inputStream))
        {
            var richText = workbook.Worksheets.First().Cell("A1").GetRichText();
            await AssertRichText(richText);
            richText.AddText(" - changed");
            workbook.SaveAs(outputStream);
        }

        using (var wb = new XLWorkbook(outputStream))
        {
            var cell = wb.Worksheets.First().Cell("A1");
            await Assert.That(cell.ShareString).IsFalse();
            await Assert.That(cell.HasRichText).IsTrue();
            var rt = cell.GetRichText();
            await Assert.That(rt.ToString()).IsEqualTo("Year (range: 3 yrs) - changed");
            await AssertRichText(rt);
        }
    }

    [Test]
    public async Task ClearInlineRichTextWhenRelevant()
    {
        using var ms = new MemoryStream();
        await TestHelper.CreateAndCompare(() =>
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet();
                var cell = ws.FirstCell();

                cell.GetRichText().AddText("Bold").SetBold().AddText(" and red").SetBold().SetFontColor(XLColor.Red);
                cell.ShareString = false;

                //wb.SaveAs(ms);
                wb.SaveAs(ms);
            }
            ms.Seek(0, SeekOrigin.Begin);

            var wb2 = new XLWorkbook(ms);
            {
                var ws = wb2.Worksheets.First();
                var cell = ws.FirstCell();

                cell.FormulaA1 = "=1 + 2";
                wb2.SaveAs(ms);
            }

            ms.Seek(0, SeekOrigin.Begin);

            return wb2;
        }, @"Other\InlinedRichText\ChangeRichTextToFormula\output.xlsx");
    }

    [Test]
    public async Task RichTextChangesContentOfItsCell()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var cell = ws.Cell(1, 1);
        var richText = cell.GetRichText();

        await Assert.That(richText.Text).IsEqualTo(cell.Value);

        richText.AddText("Hello");
        await Assert.That("Hello").IsEqualTo(cell.Value);

        var world = richText.AddText(" World");
        await Assert.That("Hello World").IsEqualTo(cell.Value);

        world.Text = " World!";
        await Assert.That("Hello World!").IsEqualTo(cell.Value);
        await Assert.That("Hello World!").IsEqualTo(cell.GetRichText().Text);

        richText.ClearText();
        await Assert.That(string.Empty).IsEqualTo(cell.Value);
    }

    [Test]
    public async Task RemovedRichTextFromCellCantBeChanged()
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var cell = ws.Cell(1, 1);
        var richText = cell.GetRichText();
        cell.Value = 4;

        await Assert.That(() => richText.AddText("Hello")).Throws<InvalidOperationException>();
    }

    [Test]
    public async Task MaintainWhitespaces()
    {
        const string textWithSpaces = "  元  気  ";
        const string phoneticsWithSpace = "  げ  ん  ";
        using var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.AddWorksheet();
            var richTextCell = ws.Cell(1, 1);
            var richText = richTextCell.GetRichText();
            richText.AddText(textWithSpaces);
            richText.Phonetics.Add(phoneticsWithSpace, 2, 3);

            wb.SaveAs(ms);
        }

        ms.Position = 0;

        using (var wb = new XLWorkbook(ms))
        {
            var ws = wb.Worksheets.First();
            var richText = ws.Cell(1, 1).GetRichText();
            await Assert.That(richText.First().Text).IsEqualTo(textWithSpaces);
            await Assert.That(richText.Phonetics.First().Text).IsEqualTo(phoneticsWithSpace);
        }
    }

    [Test]
    public async Task Empty_phonetic_run_in_shared_string_is_skipped()
    {
        // Some versions of Excel produce <rPh sb="0" eb="0"><t/></rPh> elements
        // in sharedStrings.xml for Japanese text. These have empty text and equal
        // start/end indices, which should be silently skipped rather than throwing.
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet();
        var xlCell = (XLCell)ws.Cell(1, 1);

        var sharedString = new SharedStringItem(
            new Run(new Text("日本語テスト")),
            new PhoneticRun(new Text()) { BaseTextStartIndex = 0, EndingBaseIndex = 0 },
            new PhoneticProperties { FontId = 0 });

        WorksheetSheetDataReader.SetCellText(xlCell, sharedString);

        await Assert.That(xlCell.GetRichText().Text).IsEqualTo("日本語テスト");
        await Assert.That(xlCell.GetRichText().Phonetics.Count).IsEqualTo(0);
    }

    [Test]
    public async Task Preserve_end_of_line_in_xml()
    {
        // When text run in a rich text contains end of line (regardless if CR, LF or CRLF),
        // the written element must be marked with xml:space="preserve". Excel would process
        // text differently (trim ect, see XML spec) and that means there would be a data
        // loss (trimmed ends of line). Another problem would be phonetic runs. They use indexes
        // to the text run, but if text would be trimmed, they might suddenly have out-of-bounds
        // values and Excel would try to repair the workbook.
        // The source files contains a text run with end of line at the start and end. It also
        // contains phonetic run for the kanji in the text that would be out-of-bounds if space
        // attribute there. The input is from Excel, output is by XLibur. Output must contain
        // the space attribute.
        await Assert.That(() => TestHelper.LoadSaveAndCompare(
            @"Other\RichText\kanji-with-new-line-input.xlsx",
            @"Other\RichText\kanji-with-new-line-output.xlsx")).ThrowsNothing();
    }
}
