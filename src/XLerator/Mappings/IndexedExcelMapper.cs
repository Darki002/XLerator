using System.Reflection;
using XLerator.Attributes;

namespace XLerator.Mappings;

internal class IndexedExcelMapper : ExcelMapperBase
{
    private IndexedExcelMapper() {}

    public override string? GetHeaderFor(string propertyName)
    {
        if (!PropertyIndexMap.ContainsKey(propertyName)) return null;
        
        var headerName = HeaderMap.GetValueOrDefault(propertyName, propertyName);
        return headerName;
    }

    internal static IndexedExcelMapper CreateFrom(Type type)
    {
        var mapper = new IndexedExcelMapper();

        var propertyInfos = type.GetProperties();
        foreach (var property in propertyInfos)
        {
            var attribute = property.GetCustomAttribute<ColumnIndex>();
            if (attribute is null) continue;
            mapper.PropertyIndexMap.Add(property.Name, attribute.Index);
            
            var headerAttribute = property.GetCustomAttribute<HeaderName>();
            if(headerAttribute is null) continue;
            mapper.HeaderMap.Add(property.Name, headerAttribute.Name);
        }

        return mapper;
    }
}