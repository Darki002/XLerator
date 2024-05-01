using XLerator.ExcelMappings;

namespace XLerator.Tests.ExcelMappings;

public class ExcelMapperBaseFake : ExcelMapperBase
{
    public Dictionary<string, int> PropertyIndexMapSpy => PropertyIndexMap;

    public Dictionary<string, string> HeaderMapSpy => HeaderMap;

    public void AddPropertyIndexMap(string key, int value)
    {
        PropertyIndexMap.Add(key, value);
    }
    
    public void AddHeaderMap(string key, string value)
    {
        HeaderMap.Add(key, value);
    }
}