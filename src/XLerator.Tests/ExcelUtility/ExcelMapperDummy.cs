using XLerator.Mappings;

namespace XLerator.Tests.ExcelUtility;

public class ExcelMapperDummy : ExcelMapperBase
{
    public override string? GetHeaderFor(string propertyName)
    {
        return HeaderMap[propertyName];
    }
}