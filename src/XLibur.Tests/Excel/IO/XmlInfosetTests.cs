using System.IO;
using System.Text;
using System.Threading.Tasks;
using XLibur.Excel.IO;

namespace XLibur.Tests.Excel.IO;

public class XmlInfosetTests
{
    private const string Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string Mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";

    [Test]
    [Arguments("prefix",
        $"""<worksheet xmlns="{Main}"><sheetData/></worksheet>""",
        $"""<x:worksheet xmlns:x="{Main}"><x:sheetData/></x:worksheet>""")]
    [Arguments("declaration and byte order mark",
        "﻿" + $"""<?xml version='1.0' encoding='UTF-8'?><worksheet xmlns="{Main}"/>""",
        $"""<?xml version="1.0" encoding="utf-8" standalone="yes"?><worksheet xmlns="{Main}"/>""")]
    [Arguments("attribute order",
        $"""<worksheet xmlns="{Main}"><col min="1" max="2"/></worksheet>""",
        $"""<worksheet xmlns="{Main}"><col max="2" min="1"/></worksheet>""")]
    [Arguments("unused declaration",
        $"""<worksheet xmlns="{Main}" xmlns:r="urn:r"/>""",
        $"""<worksheet xmlns="{Main}"/>""")]
    [Arguments("empty element",
        $"""<worksheet xmlns="{Main}"><sheetData/></worksheet>""",
        $"""<worksheet xmlns="{Main}"><sheetData></sheetData></worksheet>""")]
    [Arguments("whitespace between elements",
        $"""<worksheet xmlns="{Main}">""" + "\r\n  <sheetData/>\r\n</worksheet>",
        $"""<worksheet xmlns="{Main}"><sheetData/></worksheet>""")]
    [Arguments("comment",
        $"""<worksheet xmlns="{Main}"><!-- note --><sheetData/></worksheet>""",
        $"""<worksheet xmlns="{Main}"><sheetData/></worksheet>""")]
    [Arguments("CDATA",
        $"""<worksheet xmlns="{Main}"><oddHeader>a<![CDATA[&b]]></oddHeader></worksheet>""",
        $"""<worksheet xmlns="{Main}"><oddHeader>a&amp;b</oddHeader></worksheet>""")]
    [Arguments("ignorable prefix bound alike",
        $"""<worksheet xmlns="{Main}" xmlns:mc="{Mc}" xmlns:a="urn:a" mc:Ignorable="a"/>""",
        $"""<m:worksheet xmlns:m="{Main}" xmlns:q="{Mc}" xmlns:a="urn:a" q:Ignorable="a"/>""")]
    public async Task Equivalent(string difference, string left, string right)
    {
        await Assert.That(AreEquivalent(left, right)).IsTrue().Because(difference);
        await Assert.That(AreEquivalent(right, left)).IsTrue().Because(difference);
    }

    [Test]
    [Arguments("element namespace",
        $"""<worksheet xmlns="{Main}"><sheetData/></worksheet>""",
        $"""<worksheet xmlns="{Main}"><sheetData xmlns="urn:other"/></worksheet>""")]
    [Arguments("element name",
        $"""<worksheet xmlns="{Main}"><sheetData/></worksheet>""",
        $"""<worksheet xmlns="{Main}"><sheetView/></worksheet>""")]
    [Arguments("attribute value",
        $"""<worksheet xmlns="{Main}"><col min="1" max="2"/></worksheet>""",
        $"""<worksheet xmlns="{Main}"><col min="1" max="3"/></worksheet>""")]
    [Arguments("attribute namespace",
        $"""<worksheet xmlns="{Main}" xmlns:a="urn:a" xmlns:b="urn:b"><col a:id="1"/></worksheet>""",
        $"""<worksheet xmlns="{Main}" xmlns:a="urn:a" xmlns:b="urn:b"><col b:id="1"/></worksheet>""")]
    [Arguments("extra attribute",
        $"""<worksheet xmlns="{Main}"><col min="1"/></worksheet>""",
        $"""<worksheet xmlns="{Main}"><col min="1" max="1"/></worksheet>""")]
    [Arguments("extra element",
        $"""<worksheet xmlns="{Main}"><sheetData/></worksheet>""",
        $"""<worksheet xmlns="{Main}"><sheetData/><mergeCells/></worksheet>""")]
    [Arguments("child order",
        $"""<worksheet xmlns="{Main}"><a/><b/></worksheet>""",
        $"""<worksheet xmlns="{Main}"><b/><a/></worksheet>""")]
    [Arguments("nesting",
        $"""<worksheet xmlns="{Main}"><a/><b/></worksheet>""",
        $"""<worksheet xmlns="{Main}"><a><b/></a></worksheet>""")]
    [Arguments("text",
        $"""<worksheet xmlns="{Main}"><oddHeader>a</oddHeader></worksheet>""",
        $"""<worksheet xmlns="{Main}"><oddHeader>b</oddHeader></worksheet>""")]
    [Arguments("text against none",
        $"""<worksheet xmlns="{Main}"><oddHeader>a</oddHeader></worksheet>""",
        $"""<worksheet xmlns="{Main}"><oddHeader/></worksheet>""")]
    [Arguments("ignorable prefix bound otherwise",
        $"""<worksheet xmlns="{Main}" xmlns:mc="{Mc}" xmlns:a="urn:a" mc:Ignorable="a"/>""",
        $"""<worksheet xmlns="{Main}" xmlns:mc="{Mc}" xmlns:a="urn:b" mc:Ignorable="a"/>""")]
    [Arguments("required prefix bound otherwise",
        $"""<worksheet xmlns="{Main}" xmlns:mc="{Mc}" xmlns:a="urn:a"><mc:AlternateContent><mc:Choice Requires="a"/></mc:AlternateContent></worksheet>""",
        $"""<worksheet xmlns="{Main}" xmlns:mc="{Mc}"><mc:AlternateContent xmlns:a="urn:b"><mc:Choice Requires="a"/></mc:AlternateContent></worksheet>""")]
    public async Task Different(string difference, string left, string right)
    {
        await Assert.That(AreEquivalent(left, right)).IsFalse().Because(difference);
        await Assert.That(AreEquivalent(right, left)).IsFalse().Because(difference);
    }

    private static bool AreEquivalent(string left, string right)
    {
        using var l = new MemoryStream(Encoding.UTF8.GetBytes(left));
        using var r = new MemoryStream(Encoding.UTF8.GetBytes(right));
        return XmlInfoset.AreEquivalent(l, r);
    }
}
