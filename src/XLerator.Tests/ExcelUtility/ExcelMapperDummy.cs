using XLerator.Mappings;

namespace XLerator.Tests.ExcelUtility;

internal class ExcelMapperDummy : ExcelMapperBase
{
    public override string? GetHeaderFor(string propertyName)
    {
        return HeaderMap[propertyName];
    }
}