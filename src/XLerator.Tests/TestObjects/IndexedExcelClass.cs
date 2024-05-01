using XLerator.Tags;

namespace XLerator.Tests.TestObjects;

[IndexedExcel]
public class IndexedExcelClass
{
    [ExcelIndex(1)] public int Index { get; init; } = 0;

    [ExcelIndex(2)] public string Name { get; init; } = "Test";
}