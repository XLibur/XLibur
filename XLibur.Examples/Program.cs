using System;
using System.IO;
using XLibur.Examples.Creating;
using XLibur.Examples.Loading;

namespace XLibur.Examples;

public static class Program
{
    public static string BaseCreatedDirectory
    {
        get
        {
            var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Created");
            if (!Directory.Exists(path)) Directory.CreateDirectory(path);
            return path;
        }
    }

    public static string BaseModifiedDirectory
    {
        get
        {
            var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Modified");
            if (!Directory.Exists(path)) Directory.CreateDirectory(path);
            return path;
        }
    }

    private static void Main()
    {
        CreateFiles.CreateAllFiles();
        LoadFiles.LoadAllFiles();
    }
}
