using System.ComponentModel;
using DocumentFormat.OpenXml.Spreadsheet;
using XLerator.Mappings;

namespace XLerator.ExcelUtility.ExcelReading;

internal static class Helper
{
    internal static object? GetDefaultValue(Type type)
    {
        if (type == typeof(string)) return string.Empty;
        return type.IsValueType ? Activator.CreateInstance(type) : null;
    }

    internal static string? GetCellValue<T>(IReadOnlyList<Cell> cells, ExcelMapperBase excelMapper, string propertyName)
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