using XLibur.Examples.Comments;
using System.Threading.Tasks;

namespace XLibur.Tests.Examples;

public class CommentsTests
{
    [Test]
    // Windows-only: VML drawing comparison is platform-dependent: XDocument serialization produces different XML formatting on Linux vs Windows
    [RunOn(TUnit.Core.Enums.OS.Windows)]
    public async Task AddingComments()
    {
        await TestHelper.RunTestExample<AddingComments>(@"Comments\AddingComments.xlsx");
    }
}
