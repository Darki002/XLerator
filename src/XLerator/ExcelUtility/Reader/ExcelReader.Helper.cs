using System.ComponentModel;
using DocumentFormat.OpenXml.Spreadsheet;

namespace XLerator.ExcelUtility.Reader;

internal partial class ExcelReader<T>
{
    internal static object? GetDefaultValue(Type type)
    {
        if (type == typeof(string)) return string.Empty;
        return type.IsValueType ? Activator.CreateInstance(type) : null;
    }

    internal string? GetCellValue(IReadOnlyList<Cell> cells, string propertyName)
    {
        var cellIndex = excelMapper.GetColumnIndexFor(propertyName);
        ThrowHelper.ThrowIfNull(cellIndex, $"Excel file does not Match expected pattern of Type {typeof(T)}");
        return cells[(int)cellIndex! - 1].CellValue?.InnerText;
    }

    internal static object? GetValueOrDefault(Type type, string? valueString)
    {
        if (valueString is null) return GetDefaultValue(type);
        var converter = TypeDescriptor.GetConverter(type);
        return converter.ConvertFromString(valueString);
    }
}