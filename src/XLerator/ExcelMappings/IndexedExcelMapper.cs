using System.Reflection;
using XLerator.Tags;

namespace XLerator.ExcelMappings;

internal class IndexedExcelMapper : ExcelMapperBase
{
    private readonly Dictionary<string, int> propertyIndexMap;

    private readonly Dictionary<string, string> headerMap;

    private IndexedExcelMapper()
    {
        propertyIndexMap = new Dictionary<string, int>();
        headerMap = new Dictionary<string, string>();
    }
    
    internal static IndexedExcelMapper CreateFrom(Type type)
    {
        var mapper = new IndexedExcelMapper();

        var propertyInfos = type.GetProperties();
        foreach (var property in propertyInfos)
        {
            var attribute = property.GetCustomAttribute<ExcelIndex>();
            if (attribute is null) continue;
            mapper.propertyIndexMap.Add(property.Name, attribute.Index);
            
            var headerAttribute = property.GetCustomAttribute<ExcelHeaderName>();
            if(headerAttribute is null) continue;
            mapper.headerMap.Add(property.Name, headerAttribute.Name);
        }

        return mapper;
    }

    public override string GetColumnFor(string propertyName)
    {
        var columnNumber = propertyIndexMap[propertyName];
        return IntoToColumnString(columnNumber);
    }

    public override string GetHeaderNameFor(string propertyName)
    {
        return headerMap.GetValueOrDefault(propertyName, propertyName);
    }
}