using XLerator.Attributes;

namespace XLerator.Tests.TestObjects;

[IndexedExcel]
public class IndexedExcelClass
{
    [ExcelIndex(1)] 
    [ExcelHeaderName("Index")]
    public int Id { get; init; } = 0;

    [ExcelIndex(2)] public string Name { get; init; } = "Test";
}