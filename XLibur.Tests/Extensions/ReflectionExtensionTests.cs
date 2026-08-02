using XLibur.Extensions;
using System;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;

namespace XLibur.Tests.Extensions;

public class ReflectionExtensionTests
{
#pragma warning disable CS0067 // Event is never used — members exist for reflection tests
    private class TestClass
    {
        static TestClass()
        {
        }

        public static int StaticProperty { get; set; }
        public static int StaticField = 0;

        public static event EventHandler<EventArgs> StaticEvent;

        public static void StaticMethod()
        {
        }

        public const int Const = 100;

        public int InstanceProperty { get; set; }
        public int InstanceField = 0;

        public event EventHandler<EventArgs>? InstanceEvent;

#pragma warning disable CA1822 // Intentionally non-static: test verifies IsStatic() returns false
        public void InstanceMethod()
        {
        }
#pragma warning restore CA1822
    }
#pragma warning restore CS0067

    [Test]
    [Arguments(nameof(TestClass.StaticProperty), true)]
    [Arguments(nameof(TestClass.StaticField), true)]
    [Arguments(nameof(TestClass.StaticEvent), true)]
    [Arguments(nameof(TestClass.StaticMethod), true)]
    [Arguments(nameof(TestClass.Const), true)]
    [Arguments(nameof(TestClass.InstanceProperty), false)]
    [Arguments(nameof(TestClass.InstanceField), false)]
    [Arguments(nameof(TestClass.InstanceEvent), false)]
    [Arguments(nameof(TestClass.InstanceMethod), false)]
    public async Task IsStatic(string memberName, bool expectedIsStatic)
    {
        var member = typeof(TestClass).GetMember(memberName).Single();
        await Assert.That(member.IsStatic()).IsEqualTo(expectedIsStatic);
    }

    [Test]
    [Arguments(BindingFlags.Static | BindingFlags.NonPublic, true)]
    [Arguments(BindingFlags.Instance | BindingFlags.Public, false)]
    public async Task ConstructorIsStatic(BindingFlags flag, bool expectedIsStatic)
    {
        var constructors = typeof(TestClass).GetConstructors(flag);
        await Assert.That(constructors.Single().IsStatic()).IsEqualTo(expectedIsStatic);
    }
}
