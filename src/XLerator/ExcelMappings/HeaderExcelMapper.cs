using System.Reflection;
using XLerator.Tags;

namespace XLerator.ExcelMappings;

public class HeaderExcelMapper : ExcelMapperBase
{
    private HeaderExcelMapper() {}
    
    public static HeaderExcelMapper CreateFrom(Type type)
    {
        var mapper = new HeaderExcelMapper();
        
        var propertyInfos = type.GetProperties();

        var index = 0;
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

    public static HeaderExcelMapper CreateFrom(Type type, string filePath)
    {
        throw new NotImplementedException();
    }
}