using System.IO;
using XLibur.Excel;

namespace XLibur.Examples.Loading;

public static class LoadFiles
{
    public static void LoadAllFiles()
    {
        foreach (var file in Directory.GetFiles(Program.BaseCreatedDirectory))
        {
            var ext = Path.GetExtension(file).ToLowerInvariant();
            if (ext is not ".xlsx" and not ".xlsm" and not ".xltx" and not ".xltm")
                continue;

            var fileName = Path.GetFileName(file);
            LoadAndSaveFile(Path.Combine(Program.BaseCreatedDirectory, fileName), Path.Combine(Program.BaseModifiedDirectory, fileName));
        }
    }

    private static void LoadAndSaveFile(string input, string output)
    {
        var wb = new XLWorkbook(input);
        wb.SaveAs(output);
        wb.SaveAs(output);
    }
}
