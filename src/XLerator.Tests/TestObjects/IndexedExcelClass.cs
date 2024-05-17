using XLerator.Attributes;

namespace XLerator.Tests.TestObjects;

public class IndexedExcelClass
{
    [ColumnIndex(1)] 
    [HeaderName("Index")]
    public int Id { get; init; } = 0;

    [ColumnIndex(2)] public string Name { get; init; } = "Test";
}