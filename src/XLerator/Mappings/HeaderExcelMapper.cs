using System.Reflection;
using XLerator.Attributes;

namespace XLerator.ExcelMappings;

internal class HeaderExcelMapper : ExcelMapperBase
{
    private HeaderExcelMapper() {}
    
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

    internal static HeaderExcelMapper CreateFromExistingExcel(Type type, string filePath)
    {
        throw new NotImplementedException();
    }
}