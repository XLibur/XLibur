using System;

namespace XLibur.Extensions;

internal static class TypeExtensions
{
    extension(Type type)
    {
        public Type GetUnderlyingType()
        {
            return Nullable.GetUnderlyingType(type) ?? type;
        }

        public bool IsNullableType()
        {
            return Nullable.GetUnderlyingType(type) != null;
        }

        public bool IsNumber()
        {
            return type == typeof(sbyte)
                   || type == typeof(byte)
                   || type == typeof(short)
                   || type == typeof(ushort)
                   || type == typeof(int)
                   || type == typeof(uint)
                   || type == typeof(long)
                   || type == typeof(ulong)
                   || type == typeof(float)
                   || type == typeof(double)
                   || type == typeof(decimal);
        }

        public bool IsSimpleType()
        {
            return type.IsPrimitive
                   || type == typeof(string)
                   || type == typeof(DateTime)
                   || type == typeof(TimeSpan)
                   || type.IsNumber();
        }
    }
}
