using System.ComponentModel;
using DocumentFormat.OpenXml.Spreadsheet;
using XLerator.Mappings;

namespace XLerator.ExcelUtility.ExcelReading;

internal static class Helper
{
    public static T DeserializerFrom<T>(List<Cell> cells, ExcelMapperBase excelMapper)
    {
        var instanceType = typeof(T);
        var properties = instanceType.GetProperties();
        var instance = (T)Activator.CreateInstance(instanceType)!;

        foreach (var propertyInfo in properties)
        {
            var type = propertyInfo.PropertyType;
            var valueString = GetCellValue<T>(cells, excelMapper, propertyInfo.Name);

            propertyInfo.SetValue(instance, GetValueOrDefault(type, valueString));
        }

        return instance;
    }
    
    internal static object? GetDefaultValue(Type type)
    {
        if (type == typeof(string)) return string.Empty;
        return type.IsValueType ? Activator.CreateInstance(type) : null;
    }

    internal static string? GetCellValue<T>(IReadOnlyList<Cell> cells, ExcelMapperBase excelMapper, string propertyName)
    {
        var cellIndex = excelMapper.GetColumnIndexFor(propertyName);
        ThrowHelper.ThrowIfNull(cellIndex, $"Excel file does not Match expected pattern of Type {typeof(T)}");
        var cell = cells[(int)cellIndex! - 1];

        if (cell.DataType != null && cell.DataType == CellValues.SharedString)
        {
            // TODO: load shared string
        }
        
        return cell.InnerText;
    }

    internal static object? GetValueOrDefault(Type type, string? valueString)
    {
        if (valueString is null) return GetDefaultValue(type);
        var converter = TypeDescriptor.GetConverter(type);
        return converter.ConvertFromString(valueString);
    }
}