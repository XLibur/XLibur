using XLibur.Examples.Delete;
using System.IO;

namespace XLibur.Examples;

public static class ModifyFiles
{
    public static void Run()
    {
        var path = Program.BaseModifiedDirectory;
        new DeleteRows().Create(Path.Combine(path, "DeleteRows.xlsx"));
    }
}
