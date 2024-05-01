using XLerator.Tags;

namespace XLerator.Tests.TestObjects;

public class HeaderedExcelClass
{
    [ExcelHeaderName("Index")]
    public int Id { get; init; } = 0;

    public string Name { get; init; } = "Test";
}