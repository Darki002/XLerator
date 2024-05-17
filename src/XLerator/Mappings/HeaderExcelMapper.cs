using System.Reflection;
using XLerator.Attributes;

namespace XLerator.Mappings;

internal class HeaderExcelMapper : ExcelMapperBase
{
    private HeaderExcelMapper() {}

    public override string? GetHeaderFor(string propertyName)
    {
        return HeaderMap[propertyName];
    }

    internal static HeaderExcelMapper CreateFrom(Type type)
    {
        var mapper = new HeaderExcelMapper();
        
        var propertyInfos = type.GetProperties();

        var index = 1;
        foreach (var property in propertyInfos)
        {
            var headerAttribute = property.GetCustomAttribute<ExcelHeaderName>();
            var name = headerAttribute?.Name ?? property.Name;
            
            mapper.HeaderMap.Add(property.Name, name);
            mapper.PropertyIndexMap.Add(property.Name, index);
            
            index++;
        }
        
        return mapper;
    }
}