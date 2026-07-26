using System.Xml;
using XLibur.Excel.RichText;
using XLibur.Extensions;
using XLibur.Utils;
using static XLibur.Excel.XLWorkbook;
using static XLibur.Excel.IO.OpenXmlConst;

namespace XLibur.Excel.IO;

internal static class TextSerializer
{
    internal static void WriteRichTextElements(XmlWriter w, XLImmutableRichText richText, SaveContext context)
    {
        if (richText.Runs.Count == 0)
        {
            // Plain text carrying only a phonetic guide - it never had runs, so write it back as a
            // bare <t> rather than wrapping it in a run with invented formatting. This text was
            // decoded on the way in (unlike run text, which is not), so it has to be re-encoded:
            // a decoded _xHHHH_ escape is a raw control character that is not valid XML content.
            if (richText.Text.Length > 0)
                WriteText(w, XmlEncoder.EncodeString(richText.Text));
        }
        else
        {
            foreach (var textRun in richText.Runs)
            {
                var text = richText.GetRunText(textRun);
                if (text.Length > 0)
                {
                    WriteRun(w, text, textRun.Font);
                }
            }
        }

        if (richText.PhoneticsProperties is not null)
        {
            var phoneticsProps = richText.PhoneticsProperties.Value;
            foreach (var p in richText.PhoneticRuns)
            {
                w.WriteStartElement("rPh", Main2006SsNs);
                w.WriteAttribute("sb", p.StartIndex);
                w.WriteAttribute("eb", p.EndIndex);

                w.WriteStartElement("t", Main2006SsNs);
                if (p.Text.PreserveSpaces())
                    w.WritePreserveSpaceAttr();

                w.WriteString(p.Text);
                w.WriteEndElement(); // t
                w.WriteEndElement(); // rPh
            }

            var font = phoneticsProps.Font;
            if (!context.SharedFonts.TryGetValue(font, out FontInfo fi))
            {
                fi = new FontInfo { Font = font };
                context.SharedFonts.Add(font, fi);
            }

            w.WriteStartElement("phoneticPr", Main2006SsNs);
            w.WriteAttribute("fontId", fi.FontId);

            if (phoneticsProps.Alignment != XLPhoneticAlignment.Left)
                w.WriteAttributeString("alignment", phoneticsProps.Alignment.ToOpenXmlString());

            if (phoneticsProps.Type != XLPhoneticType.FullWidthKatakana)
                w.WriteAttributeString("type", phoneticsProps.Type.ToOpenXmlString());

            w.WriteEndElement(); // phoneticPr
        }
    }

    internal static void WriteRun(XmlWriter w, XLImmutableRichText richText, XLImmutableRichText.RichTextRun run)
    {
        var runText = richText.GetRunText(run);
        WriteRun(w, runText, run.Font);
    }

    private static void WriteRun(XmlWriter w, string text, XLFontValue font)
    {
        w.WriteStartElement("r", Main2006SsNs);
        w.WriteStartElement("rPr", Main2006SsNs);

        if (font.Bold)
            w.WriteEmptyElement("b");

        if (font.Italic)
            w.WriteEmptyElement("i");

        if (font.Strikethrough)
            w.WriteEmptyElement("strike");

        // Three attributes are not stored/written:
        // * outline - doesn't do anything and likely only works in Word.
        // * condense - legacy compatibility setting for macs
        // * extend - legacy compatibility setting for pre-xlsx Excels
        // None have sensible descriptions.

        if (font.Shadow)
            w.WriteEmptyElement("shadow");

        if (font.Underline != XLFontUnderlineValues.None)
            WriteRunProperty(w, "u", font.Underline.ToOpenXmlString());

        WriteRunProperty(w, "vertAlign", font.VerticalAlignment.ToOpenXmlString());
        WriteRunProperty(w, "sz", font.FontSize);

        // An unset color means the run is automatic - Excel resolves it against the theme, and
        // conditional formatting can still override it. Writing an explicit black here would pin
        // it down permanently.
        if (!XLColor.IsUnset(font.FontColor.Key))
            w.WriteColor("color", font.FontColor);

        WriteRunProperty(w, "rFont", font.FontName);
        WriteRunProperty(w, "family", (int)font.FontFamilyNumbering);

        if (font.FontCharSet != XLFontCharSet.Default)
            WriteRunProperty(w, "charset", (int)font.FontCharSet);

        if (font.FontScheme != XLFontScheme.None)
            WriteRunProperty(w, "scheme", font.FontScheme.ToOpenXml());

        w.WriteEndElement(); // rPr

        WriteText(w, text);

        w.WriteEndElement(); // r
    }

    private static void WriteText(XmlWriter w, string text)
    {
        w.WriteStartElement("t", Main2006SsNs);
        if (text.PreserveSpaces())
            w.WritePreserveSpaceAttr();

        w.WriteString(text);
        w.WriteEndElement(); // t
    }

    private static void WriteRunProperty(XmlWriter w, string elName, string val)
    {
        w.WriteStartElement(elName, Main2006SsNs);
        w.WriteAttributeString("val", val);
        w.WriteEndElement();
    }

    private static void WriteRunProperty(XmlWriter w, string elName, int val)
    {
        w.WriteStartElement(elName, Main2006SsNs);
        w.WriteAttribute("val", val);
        w.WriteEndElement();
    }

    private static void WriteRunProperty(XmlWriter w, string elName, double val)
    {
        w.WriteStartElement(elName, Main2006SsNs);
        w.WriteAttribute("val", val);
        w.WriteEndElement();
    }
}
