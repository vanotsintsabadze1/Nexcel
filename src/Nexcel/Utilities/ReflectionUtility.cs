using System.Reflection;

namespace Nexcel.Utilities;

internal static class ReflectionUtility
{
    public static IEnumerable<string> ExtractPropertyNames(Type type)
    {
        var names = new List<string>();

        var properties = type.GetProperties(BindingFlags.Public | BindingFlags.Instance)
                             .Where(p => p.DeclaringType == type);

        foreach (var property in properties)
        {
            names.Add(property.Name);
        }

        return names;
    }
}
