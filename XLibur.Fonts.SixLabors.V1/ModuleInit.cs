using System.Runtime.CompilerServices;

namespace XLibur.Fonts.SixLabors.V1;

internal static class ModuleInit
{
    [ModuleInitializer]
    internal static void Initialize() => SixLaborsV1FontBootstrap.Register();
}
