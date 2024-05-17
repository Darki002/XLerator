using XLerator.Attributes;

namespace XLerator.Tests.TestObjects;

[NoExcelIndexes]
public class HeaderedExcelClass
{
    [HeaderName("Index")]
    public int Id { get; init; } = 0;

    public string Name { get; init; } = "Test";
}