using System.Reflection;
using XLibur.Excel;
using XLibur.Excel.Drawings;

namespace XLibur.Examples.ImageHandling;

public class ImageFormats : IXLExample
{
    public void Create(string filePath)
    {
        var wb = new XLWorkbook();
        IXLWorksheet ws;

        using (var fs = Assembly.GetExecutingAssembly()
                   .GetManifestResourceStream("XLibur.Examples.Resources.ImageHandling.jpg"))
        {
            #region Jpeg

            ws = wb.Worksheets.Add("Jpg");
            ws.AddPicture(fs, XLPictureFormat.Jpeg, "JpegImage")
                .MoveTo(ws.Cell(1, 1));

            #endregion Jpeg
        }

        using (var fs = Assembly.GetExecutingAssembly()
                   .GetManifestResourceStream("XLibur.Examples.Resources.ImageHandling.png"))
        {
            #region Png

            ws = wb.Worksheets.Add("Png");
            ws.AddPicture(fs, XLPictureFormat.Png, "PngImage")
                .MoveTo(ws.Cell(1, 1));

            #endregion Png

            wb.SaveAs(filePath);
        }
    }
}
