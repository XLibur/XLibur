using System.ComponentModel;
using System.Globalization;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.Json;
using XLibur.Operations;

namespace XLibur.Cli;

internal static class RequestBinder
{
    private static readonly Dictionary<(Type Request, string Option), (string Property, Func<string, object> Parse)> Shorthands = new()
    {
        [(typeof(EditRequest), "set")] = (nameof(EditRequest.Edits), CellEdit.Parse),
    };

    internal static object Bind(IOperation operation, IReadOnlyList<string> args)
    {
        var binding = new Binding(operation);
        var queue = new Queue<string>(args);
        while (queue.TryDequeue(out var arg))
        {
            if (!arg.StartsWith("--", StringComparison.Ordinal) || arg == "--")
                binding.Positional(arg);
            else
                binding.Option(arg, queue);
        }

        return binding.Build();
    }

    private sealed class Binding(IOperation operation)
    {
        private readonly Type _type = operation.RequestType;
        private readonly List<PropertyInfo> _properties = Properties(operation.RequestType);
        private readonly Dictionary<PropertyInfo, object?> _values = [];
        private readonly Dictionary<PropertyInfo, List<object>> _lists = [];
        private int _positional;

        public void Positional(string arg)
        {
            var target = _properties.FirstOrDefault(p => p.GetCustomAttribute<ArgumentAttribute>()?.Position == _positional)
                         ?? throw Usage($"Unexpected argument '{arg}'.");
            _values[target] = Convert(target.PropertyType, arg, target);
            _positional++;
        }

        public void Option(string arg, Queue<string> rest)
        {
            var (name, inline) = SplitOption(arg[2..]);
            if (Shorthands.TryGetValue((_type, name), out var shorthand))
            {
                Add(_properties.Single(p => p.Name == shorthand.Property), shorthand.Parse(inline ?? Next(rest, arg)));
                return;
            }

            var negated = name.StartsWith("no-", StringComparison.Ordinal);
            var property = Find(_properties, name) ?? (negated ? Find(_properties, name[3..]) : null)
                           ?? throw Usage($"Unknown option '--{name}' for {operation.Name}.");

            if (Unwrap(property.PropertyType) == typeof(bool))
                _values[property] = Flag(!negated, inline, arg, rest);
            else if (ElementType(property.PropertyType) == typeof(string))
                Add(property, inline ?? Next(rest, arg));
            else
                _values[property] = Convert(property.PropertyType, inline ?? Next(rest, arg), property);
        }

        public object Build()
        {
            foreach (var (property, items) in _lists)
            {
                var existing = _values.TryGetValue(property, out var parsed) && parsed is System.Collections.IEnumerable e
                    ? e.Cast<object>()
                    : [];
                _values[property] = MakeList(property.PropertyType, existing.Concat(items));
            }

            if (_properties.Where(IsRequired).FirstOrDefault(p => !_values.ContainsKey(p)) is { } missing)
                throw Usage($"{operation.Name} needs {Spelling(missing)}.");

            var request = Activator.CreateInstance(_type)!;
            foreach (var (property, value) in _values)
                property.SetValue(request, value);
            return request;
        }

        private static bool Flag(bool bareValue, string? inline, string arg, Queue<string> rest)
        {
            if (inline is not null)
                return ParseBool(inline, arg);
            if (rest.TryPeek(out var next) && bool.TryParse(next, out var explicitValue))
            {
                rest.Dequeue();
                return explicitValue;
            }

            return bareValue;
        }

        private void Add(PropertyInfo property, object item)
        {
            if (!_lists.TryGetValue(property, out var list))
                _lists[property] = list = [];
            list.Add(item);
        }
    }

    internal static string Help(IOperation operation)
    {
        var text = new StringBuilder();
        var properties = Properties(operation.RequestType);
        var positionals = properties.Where(p => p.GetCustomAttribute<ArgumentAttribute>() is not null)
            .OrderBy(p => p.GetCustomAttribute<ArgumentAttribute>()!.Position).ToList();
        text.Append("xlibur ").Append(Command(operation));
        foreach (var p in positionals)
            text.Append(IsRequired(p) ? $" <{Kebab(p.Name)}>" : $" [{Kebab(p.Name)}]");
        text.Append(" [options]\n\n").Append(operation.Description).Append("\n\n");

        foreach (var p in properties)
        {
            var spelling = p.GetCustomAttribute<ArgumentAttribute>() is not null ? $"<{Kebab(p.Name)}>" : "--" + Kebab(p.Name);
            text.Append("  ").Append(spelling.PadRight(18)).Append(' ')
                .Append(p.GetCustomAttribute<DescriptionAttribute>()?.Description).Append('\n');
        }

        foreach (var shorthand in Shorthands.Keys.Where(k => k.Request == operation.RequestType))
            text.Append("  --").Append(shorthand.Option.PadRight(16)).Append(" Shorthand edit, repeatable: B2=42 sets a value, B2==SUM(A1:A9) a formula, B2= clears.\n");
        text.Append("  --json             Print the result as JSON.\n\n");
        text.Append("A flag is --name, --name false or --no-name. A list option repeats. An option holding objects takes JSON.\n");
        return text.ToString();
    }

    internal static string Command(IOperation operation) => operation.Name.Replace('_', '-');

    private static List<PropertyInfo> Properties(Type type)
        => type.GetProperties(BindingFlags.Public | BindingFlags.Instance).Where(p => p.CanWrite && p.Name != "EqualityContract").ToList();

    private static PropertyInfo? Find(List<PropertyInfo> properties, string kebab)
        => properties.FirstOrDefault(p => Kebab(p.Name) == kebab);

    private static bool IsRequired(PropertyInfo property) => property.GetCustomAttribute<RequiredMemberAttribute>() is not null;

    private static string Spelling(PropertyInfo property)
        => property.GetCustomAttribute<ArgumentAttribute>() is not null ? $"<{Kebab(property.Name)}>" : "--" + Kebab(property.Name);

    private static string Kebab(string name)
    {
        var text = new StringBuilder();
        foreach (var c in name)
        {
            if (char.IsUpper(c) && text.Length > 0)
                text.Append('-');
            text.Append(char.ToLowerInvariant(c));
        }

        return text.ToString();
    }

    private static (string Name, string? Inline) SplitOption(string option)
    {
        var equals = option.IndexOf('=');
        return equals < 0 ? (option, null) : (option[..equals], option[(equals + 1)..]);
    }

    private static string Next(Queue<string> rest, string option)
        => rest.TryDequeue(out var value) ? value : throw Usage($"{option} needs a value.");

    private static bool ParseBool(string text, string option)
        => bool.TryParse(text, out var value) ? value : throw Usage($"{option} takes true or false.");

    private static Type Unwrap(Type type) => Nullable.GetUnderlyingType(type) ?? type;

    private static Type? ElementType(Type type)
        => type.IsGenericType && type.GetGenericTypeDefinition() == typeof(IReadOnlyList<>) ? type.GetGenericArguments()[0] : null;

    private static object MakeList(Type listType, IEnumerable<object> items)
    {
        var element = ElementType(listType)!;
        var list = (System.Collections.IList)Activator.CreateInstance(typeof(List<>).MakeGenericType(element))!;
        foreach (var item in items)
            list.Add(item);
        return list;
    }

    private static object? Convert(Type type, string raw, PropertyInfo property)
    {
        var target = Unwrap(type);
        try
        {
            if (target == typeof(string))
                return raw;
            if (target == typeof(int))
                return int.Parse(raw, CultureInfo.InvariantCulture);
            if (target == typeof(double))
                return double.Parse(raw, CultureInfo.InvariantCulture);
            if (target.IsEnum)
                return Enum.Parse(target, raw, ignoreCase: true);
            return JsonSerializer.Deserialize(raw, type, Rendering.JsonOptions);
        }
        catch (Exception e) when (e is FormatException or OverflowException or ArgumentException or JsonException)
        {
            throw Usage($"{Spelling(property)}: '{raw}' is not valid ({e.Message}).");
        }
    }

    private static OperationException Usage(string message) => OperationException.BadInput("usage", message);
}
