namespace XLerator.ExcelUtility.Reader;

public static class Helper
{
    public static object? GetDefaultValue(Type type)
    {
        if (type == typeof(string)) return string.Empty;
        return type.IsValueType ? Activator.CreateInstance(type) : null;
    }
}