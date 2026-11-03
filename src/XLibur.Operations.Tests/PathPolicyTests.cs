using System.Runtime.InteropServices;
using System.Text;
using XLibur.Operations.Paths;

namespace XLibur.Operations.Tests;

public class PathPolicyTests
{
    [Test]
    public async Task A_file_inside_the_root_is_read_by_its_canonical_path()
    {
        using var root = new TempDirectory();
        var file = root.File("book.xlsx");
        var policy = PathPolicy.Restricted([root.Path], allowWrite: false);

        using var opened = policy.OpenRead(file);

        await Assert.That(PathPolicy.SamePath(opened.FullPath, Path.GetFullPath(file))).IsTrue();
    }

    [Test]
    public async Task A_relative_path_is_relative_to_the_first_root_not_the_working_directory()
    {
        using var root = new TempDirectory();
        root.File(Path.Combine("sub", "book.xlsx"));
        var policy = PathPolicy.Restricted([root.Path], allowWrite: false);

        using var opened = policy.OpenRead(Path.Combine("sub", "book.xlsx"));

        await Assert.That(opened.FullPath.EndsWith(Path.Combine("sub", "book.xlsx"), StringComparison.OrdinalIgnoreCase)).IsTrue();
    }

    [Test]
    public async Task Dot_dot_out_of_the_root_is_refused()
    {
        using var outside = new TempDirectory();
        using var root = new TempDirectory();
        var secret = outside.File("secret.xlsx");
        var policy = PathPolicy.Restricted([root.Path], allowWrite: false);
        var escape = Path.Combine(root.Path, "..", Path.GetFileName(outside.Path), "secret.xlsx");

        await Assert.That(File.Exists(escape)).IsTrue();
        await Expect.Refused(() => policy.OpenRead(escape), ErrorKind.Forbidden, "path_outside_roots");
        await Expect.Refused(() => policy.OpenRead(secret), ErrorKind.Forbidden, "path_outside_roots");
    }

    [Test]
    public async Task A_sibling_whose_name_starts_with_the_root_name_is_outside()
    {
        using var root = new TempDirectory();
        var sibling = root.Path + "-evil";
        Directory.CreateDirectory(sibling);
        try
        {
            File.WriteAllText(Path.Combine(sibling, "book.xlsx"), "x");
            var policy = PathPolicy.Restricted([root.Path], allowWrite: false);

            await Expect.Refused(() => policy.OpenRead(Path.Combine(sibling, "book.xlsx")), ErrorKind.Forbidden, "path_outside_roots");
        }
        finally
        {
            Directory.Delete(sibling, recursive: true);
        }
    }

    [Test]
    public async Task A_case_variant_of_the_root_is_the_root_on_Windows()
    {
        Skip.Unless(OperatingSystem.IsWindows(), "Windows file names are case-insensitive; Linux names are not.");
        using var root = new TempDirectory();
        var file = root.File("Book.xlsx");
        var policy = PathPolicy.Restricted([root.Path.ToUpperInvariant()], allowWrite: false);

        using var opened = policy.OpenRead(file.ToLowerInvariant());

        await Assert.That(opened.FullPath).EndsWith("Book.xlsx");
    }

    [Test]
    [Arguments("book.xlsx:secret", "alternate_data_stream")]
    [Arguments("book.xlsx::$DATA", "alternate_data_stream")]
    [Arguments("book.xlsx.", "trailing_dot_or_space")]
    [Arguments("book.xlsx ", "trailing_dot_or_space")]
    [Arguments("dir. \\book.xlsx", "trailing_dot_or_space")]
    [Arguments("CON.xlsx", "reserved_device_name")]
    [Arguments("nul.csv", "reserved_device_name")]
    [Arguments("com1.json", "reserved_device_name")]
    [Arguments("lpt\u00B9.xlsx", "reserved_device_name")]
    [Arguments("book.exe", "extension_not_allowed")]
    [Arguments("book.xlsx.lnk", "extension_not_allowed")]
    [Arguments("noextension", "extension_not_allowed")]
    public async Task A_path_that_names_no_ordinary_file_is_refused_before_it_is_opened(string name, string code)
    {
        using var root = new TempDirectory();
        root.File("book.xlsx");
        var policy = PathPolicy.Restricted([root.Path], allowWrite: false);

        await Expect.Refused(() => policy.OpenRead(Path.Combine(root.Path, name)), ErrorKind.Forbidden, code);
    }

    [Test]
    [Arguments(@"\\attacker.example\share\book.xlsx")]
    [Arguments("//attacker.example/share/book.xlsx")]
    [Arguments(@"\\?\C:\Windows\book.xlsx")]
    [Arguments(@"\\.\C:\Windows\book.xlsx")]
    [Arguments(@"\\?\UNC\attacker.example\share\book.xlsx")]
    [Arguments(@"\??\C:\Windows\book.xlsx")]
    public async Task Network_and_device_paths_are_refused_without_being_touched(string path)
    {
        using var root = new TempDirectory();
        var policy = PathPolicy.Restricted([root.Path], allowWrite: false);

        await Expect.Refused(() => policy.OpenRead(path), ErrorKind.Forbidden, "device_or_unc_path");
    }

    [Test]
    public async Task A_NUL_character_is_refused()
    {
        using var root = new TempDirectory();
        var policy = PathPolicy.Restricted([root.Path], allowWrite: false);

        await Expect.Refused(() => policy.OpenRead("book\0.xlsx"), ErrorKind.Forbidden, "path_invalid");
    }

    [Test]
    public async Task A_junction_out_of_the_root_is_refused()
    {
        Skip.Unless(OperatingSystem.IsWindows(), "Junctions are a Windows feature.");
        using var outside = new TempDirectory();
        using var root = new TempDirectory();
        outside.File("secret.xlsx");
        Junctions.Create(root.Combine("link"), outside.Path);
        var policy = PathPolicy.Restricted([root.Path], allowWrite: true);

        await Assert.That(File.Exists(root.Combine("link", "secret.xlsx"))).IsTrue();
        await Expect.Refused(() => policy.OpenRead(root.Combine("link", "secret.xlsx")), ErrorKind.Forbidden, "path_outside_roots");
        await Expect.Refused(() => policy.ResolveWrite(root.Combine("link", "new.xlsx"), overwrite: false), ErrorKind.Forbidden, "path_outside_roots");
    }

    [Test]
    public async Task A_junction_that_stays_inside_the_root_is_followed()
    {
        Skip.Unless(OperatingSystem.IsWindows(), "Junctions are a Windows feature.");
        using var root = new TempDirectory();
        root.File(Path.Combine("data", "book.xlsx"));
        Junctions.Create(root.Combine("alias"), root.Combine("data"));
        var policy = PathPolicy.Restricted([root.Path], allowWrite: false);

        using var opened = policy.OpenRead(root.Combine("alias", "book.xlsx"));

        await Assert.That(PathPolicy.SamePath(opened.FullPath, root.Combine("data", "book.xlsx"))).IsTrue();
    }

    [Test]
    public async Task A_root_given_through_a_junction_is_its_target()
    {
        Skip.Unless(OperatingSystem.IsWindows(), "Junctions are a Windows feature.");
        using var real = new TempDirectory();
        using var holder = new TempDirectory();
        var file = real.File("book.xlsx");
        Junctions.Create(holder.Combine("root"), real.Path);
        var policy = PathPolicy.Restricted([holder.Combine("root")], allowWrite: false);

        using var direct = policy.OpenRead(file);
        using var viaLink = policy.OpenRead(holder.Combine("root", "book.xlsx"));

        await Assert.That(PathPolicy.SamePath(direct.FullPath, viaLink.FullPath)).IsTrue();
    }

    [Test]
    public async Task A_symbolic_directory_link_out_of_the_root_is_refused()
    {
        using var outside = new TempDirectory();
        using var root = new TempDirectory();
        outside.File("secret.xlsx");
        Junctions.CreateSymbolicDirectoryOrSkip(root.Combine("link"), outside.Path);
        var policy = PathPolicy.Restricted([root.Path], allowWrite: false);

        await Expect.Refused(() => policy.OpenRead(root.Combine("link", "secret.xlsx")), ErrorKind.Forbidden, "path_outside_roots");
    }

    [Test]
    public async Task A_symbolic_file_link_out_of_the_root_is_refused()
    {
        using var outside = new TempDirectory();
        using var root = new TempDirectory();
        var secret = outside.File("secret.xlsx");
        Junctions.CreateSymbolicFileOrSkip(root.Combine("innocent.xlsx"), secret);
        var policy = PathPolicy.Restricted([root.Path], allowWrite: true);

        await Expect.Refused(() => policy.OpenRead(root.Combine("innocent.xlsx")), ErrorKind.Forbidden, "path_outside_roots");
        await Expect.Refused(() => policy.ResolveWrite(root.Combine("innocent.xlsx"), overwrite: true), ErrorKind.Forbidden, "destination_is_link");
    }

    [Test]
    public async Task A_relative_symbolic_link_that_climbs_out_is_refused()
    {
        using var outside = new TempDirectory();
        using var root = new TempDirectory();
        outside.File("secret.xlsx");
        var relative = Path.Combine("..", Path.GetFileName(outside.Path));
        Junctions.CreateSymbolicDirectoryOrSkip(root.Combine("up"), relative);
        var policy = PathPolicy.Restricted([root.Path], allowWrite: false);

        await Expect.Refused(() => policy.OpenRead(root.Combine("up", "secret.xlsx")), ErrorKind.Forbidden, "path_outside_roots");
    }

    [Test]
    public async Task A_symbolic_link_to_a_network_share_is_refused_without_contacting_it()
    {
        Skip.Unless(OperatingSystem.IsWindows(), "UNC targets are a Windows feature.");
        using var root = new TempDirectory();
        Junctions.CreateSymbolicDirectoryOrSkip(root.Combine("share"), @"\\attacker.invalid\share");
        var policy = PathPolicy.Restricted([root.Path], allowWrite: false);

        await Expect.Refused(() => policy.OpenRead(root.Combine("share", "book.xlsx")), ErrorKind.Forbidden, "link_outside_roots");
    }

    [Test]
    public async Task A_short_name_inside_the_root_is_expanded_and_accepted()
    {
        Skip.Unless(OperatingSystem.IsWindows(), "8.3 short names are a Windows feature.");
        using var root = new TempDirectory();
        var file = root.File(Path.Combine("A long directory name", "book.xlsx"));
        var shortDirectory = ShortName(root.Combine("A long directory name"));
        Skip.When(shortDirectory is null || !shortDirectory.Contains('~'), "8.3 names are disabled on this volume.");
        var policy = PathPolicy.Restricted([root.Path], allowWrite: false);

        using var opened = policy.OpenRead(Path.Combine(shortDirectory!, "book.xlsx"));

        await Assert.That(PathPolicy.SamePath(opened.FullPath, file)).IsTrue();
    }

    [Test]
    public async Task A_short_name_of_a_directory_outside_the_root_is_refused()
    {
        Skip.Unless(OperatingSystem.IsWindows(), "8.3 short names are a Windows feature.");
        using var parent = new TempDirectory();
        var root = Directory.CreateDirectory(Path.Combine(parent.Path, "data")).FullName;
        var outside = Directory.CreateDirectory(Path.Combine(parent.Path, "data-secret-store")).FullName;
        File.WriteAllText(Path.Combine(outside, "secret.xlsx"), "x");
        var shortOutside = ShortName(outside);
        Skip.When(shortOutside is null || !shortOutside.Contains('~'), "8.3 names are disabled on this volume.");
        var policy = PathPolicy.Restricted([root], allowWrite: false);

        await Expect.Refused(() => policy.OpenRead(Path.Combine(shortOutside!, "secret.xlsx")), ErrorKind.Forbidden, "path_outside_roots");
    }

    [Test]
    public async Task Writes_are_refused_while_they_are_off()
    {
        using var root = new TempDirectory();
        var policy = PathPolicy.Restricted([root.Path], allowWrite: false);

        await Expect.Refused(() => policy.ResolveWrite(root.Combine("out.xlsx"), overwrite: true), ErrorKind.Forbidden, "writes_disabled");
    }

    [Test]
    public async Task A_write_outside_the_root_is_refused()
    {
        using var outside = new TempDirectory();
        using var root = new TempDirectory();
        var policy = PathPolicy.Restricted([root.Path], allowWrite: true);

        await Expect.Refused(() => policy.ResolveWrite(Path.Combine(outside.Path, "out.xlsx"), overwrite: true), ErrorKind.Forbidden, "path_outside_roots");
        await Expect.Refused(() => policy.ResolveWrite(Path.Combine(root.Path, "..", "out.xlsx"), overwrite: true), ErrorKind.Forbidden, "path_outside_roots");
    }

    [Test]
    public async Task A_write_needs_overwrite_to_replace_a_file()
    {
        using var root = new TempDirectory();
        var existing = root.File("out.csv", "old");
        var policy = PathPolicy.Restricted([root.Path], allowWrite: true);

        await Expect.Refused(() => policy.ResolveWrite(existing, overwrite: false), ErrorKind.BadInput, "file_exists");
        await Expect.Refused(() => policy.ResolveWrite(root.Combine("out.bat"), overwrite: true), ErrorKind.Forbidden, "extension_not_allowed");
    }

    [Test]
    public async Task An_atomic_write_replaces_the_file_and_leaves_no_temporary_file()
    {
        using var root = new TempDirectory();
        var existing = root.File("out.csv", "old");
        var policy = PathPolicy.Restricted([root.Path], allowWrite: true);

        var destination = policy.ResolveWrite(existing, overwrite: true);
        policy.WriteAtomically(destination, stream => stream.Write("new"u8));

        await Assert.That(File.ReadAllText(existing)).IsEqualTo("new");
        await Assert.That(Directory.GetFiles(root.Path).Length).IsEqualTo(1);
    }

    [Test]
    public async Task A_failed_write_leaves_the_destination_as_it_was()
    {
        using var root = new TempDirectory();
        var existing = root.File("out.csv", "old");
        var policy = PathPolicy.Restricted([root.Path], allowWrite: true);
        var destination = policy.ResolveWrite(existing, overwrite: true);

        await Assert.That(() => policy.WriteAtomically(destination, stream =>
        {
            stream.Write("half"u8);
            throw new InvalidOperationException("save failed");
        })).Throws<InvalidOperationException>();

        await Assert.That(File.ReadAllText(existing)).IsEqualTo("old");
        await Assert.That(Directory.GetFiles(root.Path).Length).IsEqualTo(1);
    }

    [Test]
    public async Task The_size_limit_is_checked_on_the_open_file()
    {
        using var root = new TempDirectory();
        var file = root.File("big.csv", new string('x', 2000));
        var policy = PathPolicy.Restricted([root.Path], allowWrite: false);
        using var opened = policy.OpenRead(file);

        await Expect.Refused(() => opened.ReadAllBytes(1000), ErrorKind.BadInput, "file_too_large");
    }

    [Test]
    public async Task A_missing_file_and_a_directory_are_bad_input()
    {
        using var root = new TempDirectory();
        Directory.CreateDirectory(root.Combine("folder.xlsx"));
        var policy = PathPolicy.Restricted([root.Path], allowWrite: false);

        await Expect.Refused(() => policy.OpenRead(root.Combine("missing.xlsx")), ErrorKind.BadInput, "file_not_found");
        await Expect.Refused(() => policy.OpenRead(root.Combine("folder.xlsx")), ErrorKind.BadInput, "access_denied");
    }

    [Test]
    public async Task Unrestricted_reads_anywhere_but_still_refuses_streams_and_device_names()
    {
        using var anywhere = new TempDirectory();
        var file = anywhere.File("book.xlsx");
        var policy = PathPolicy.Unrestricted();

        using (policy.OpenRead(file))
        {
        }

        await Expect.Refused(() => policy.OpenRead(file + ":stream"), ErrorKind.Forbidden, "alternate_data_stream");
        await Expect.Refused(() => policy.OpenRead(Path.Combine(anywhere.Path, "aux.xlsx")), ErrorKind.Forbidden, "reserved_device_name");
    }

    [Test]
    public async Task A_root_must_exist()
    {
        await Assert.That(() => PathPolicy.Restricted([Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"))], false))
            .Throws<DirectoryNotFoundException>();
    }

    [Test]
    public async Task A_sibling_that_differs_only_in_case_is_outside_a_case_sensitive_directory()
    {
        Skip.Unless(OperatingSystem.IsWindows(), "Per-directory case sensitivity is an NTFS feature.");
        using var parent = new TempDirectory();
        var enable = System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(
            "fsutil.exe", ["file", "setCaseSensitiveInfo", parent.Path, "enable"]) { RedirectStandardOutput = true, RedirectStandardError = true })!;
        enable.WaitForExit();
        Skip.When(enable.ExitCode != 0, "fsutil could not make the directory case-sensitive (the WSL optional feature may be off).");
        var root = Directory.CreateDirectory(Path.Combine(parent.Path, "Data")).FullName;
        var sibling = Directory.CreateDirectory(Path.Combine(parent.Path, "data")).FullName;
        File.WriteAllText(Path.Combine(sibling, "secret.xlsx"), "x");
        var policy = PathPolicy.Restricted([root], allowWrite: true);

        await Expect.Refused(() => policy.OpenRead(Path.Combine(sibling, "secret.xlsx")), ErrorKind.Forbidden, "path_outside_roots");
        await Expect.Refused(() => policy.ResolveWrite(Path.Combine(sibling, "new.xlsx"), overwrite: false), ErrorKind.Forbidden, "path_outside_roots");
    }

    [Test]
    public async Task A_sibling_whose_name_dotnet_case_folds_onto_the_root_is_outside()
    {
        using var parent = new TempDirectory();
        var root = Directory.CreateDirectory(Path.Combine(parent.Path, "dataΜ")).FullName;
        var sibling = Path.Combine(parent.Path, "dataµ");
        Skip.When(Directory.Exists(sibling), "This file system folds µ and Μ together, so the sibling cannot exist.");
        Directory.CreateDirectory(sibling);
        File.WriteAllText(Path.Combine(sibling, "secret.xlsx"), "x");
        var policy = PathPolicy.Restricted([root], allowWrite: false);

        await Expect.Refused(() => policy.OpenRead(Path.Combine(sibling, "secret.xlsx")), ErrorKind.Forbidden, "path_outside_roots");
    }

    [Test]
    public async Task A_link_with_an_allowed_name_to_a_file_with_another_extension_is_refused()
    {
        using var root = new TempDirectory();
        var secret = root.File("service.env", "{\"API_KEY\":\"x\"}");
        Junctions.CreateSymbolicFileOrSkip(root.Combine("data.json"), secret);
        var policy = PathPolicy.Restricted([root.Path], allowWrite: false);

        await Expect.Refused(() => policy.OpenRead(root.Combine("data.json")), ErrorKind.Forbidden, "extension_not_allowed");
    }

    [Test]
    public async Task A_path_outside_the_roots_is_refused_before_its_directory_is_probed()
    {
        using var root = new TempDirectory();
        var policy = PathPolicy.Restricted([root.Path], allowWrite: true);
        var missing = Path.Combine(Path.GetTempPath(), "no-such-dir-" + Guid.NewGuid().ToString("N"), "out.xlsx");

        await Expect.Refused(() => policy.ResolveWrite(missing, overwrite: false), ErrorKind.Forbidden, "path_outside_roots");
        await Expect.Refused(() => policy.OpenRead(missing), ErrorKind.Forbidden, "path_outside_roots");
    }

    [Test]
    public async Task A_drive_or_file_system_root_as_the_root_admits_paths_under_it()
    {
        using var directory = new TempDirectory();
        var file = directory.File("book.xlsx");
        var policy = PathPolicy.Restricted([Path.GetPathRoot(directory.Path)!], allowWrite: false);

        using var opened = policy.OpenRead(file);

        await Assert.That(opened.FullPath).EndsWith("book.xlsx");
    }

    [Test]
    public async Task A_very_long_path_is_refused_without_walking_it()
    {
        using var root = new TempDirectory();
        var policy = PathPolicy.Restricted([root.Path], allowWrite: false);
        var deep = Path.Combine([root.Path, .. Enumerable.Repeat("a", 100), "book.xlsx"]);

        await Expect.Refused(() => policy.OpenRead(deep), ErrorKind.BadInput, "path_too_long");
        await Expect.Refused(() => policy.OpenRead(new string('a', 2000) + ".xlsx"), ErrorKind.BadInput, "path_too_long");
    }

    private static string? ShortName(string path)
    {
        var buffer = new StringBuilder(1024);
        return GetShortPathNameW(path, buffer, buffer.Capacity) == 0 ? null : buffer.ToString();
    }

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern int GetShortPathNameW(string longPath, StringBuilder shortPath, int length);
}
