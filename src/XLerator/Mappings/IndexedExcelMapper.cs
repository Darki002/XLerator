using System.Reflection;
using XLerator.Attributes;

namespace XLerator.Mappings;

internal class IndexedExcelMapper : ExcelMapperBase
{
    private IndexedExcelMapper() {}

    public override (string, int)? GetHeaderFor(string propertyName)
    {
        if (PropertyIndexMap.TryGetValue(propertyName, out var headerIndex))
        {
            var headerName = HeaderMap.GetValueOrDefault(propertyName, propertyName);
            return (headerName, headerIndex);
        }
        return null;
    }

    internal static IndexedExcelMapper CreateFrom(Type type)
    {
        var mapper = new IndexedExcelMapper();

        var propertyInfos = type.GetProperties();
        foreach (var property in propertyInfos)
        {
            var attribute = property.GetCustomAttribute<ExcelIndex>();
            if (attribute is null) continue;
            mapper.PropertyIndexMap.Add(property.Name, attribute.Index);
            
            var headerAttribute = property.GetCustomAttribute<ExcelHeaderName>();
            if(headerAttribute is null) continue;
            mapper.HeaderMap.Add(property.Name, headerAttribute.Name);
        }

        return mapper;
    }
}