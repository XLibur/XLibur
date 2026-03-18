using System;
using System.Linq.Expressions;
using System.Reflection;

namespace XLibur;

public static class AttributeExtensions
{
    public static TAttribute[] GetAttributes<TAttribute>(
        this MemberInfo member)
        where TAttribute : Attribute
    {
        var attributes = member.GetCustomAttributes(typeof(TAttribute), true);

        return (TAttribute[])attributes;
    }

    extension<T>(T instance)
    {
        public MethodInfo GetMethod(Expression<Func<T, object>> methodSelector)
        {
            return ((MethodCallExpression)methodSelector.Body).Method;
        }

        public MethodInfo GetMethod(Expression<Action<T>> methodSelector)
        {
            return ((MethodCallExpression)methodSelector.Body).Method;
        }
    }

    public static bool HasAttribute<TAttribute>(
        this MemberInfo member)
        where TAttribute : Attribute
    {
        return member.GetAttributes<TAttribute>().Length != 0;
    }
}
