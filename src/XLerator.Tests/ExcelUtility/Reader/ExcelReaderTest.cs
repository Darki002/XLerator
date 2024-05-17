using XLerator.ExcelUtility;
using XLerator.ExcelUtility.Creator;
using XLerator.ExcelUtility.Reader;
using XLerator.Mappings;
using XLerator.Tests.TestObjects;

namespace XLerator.Tests.ExcelUtility.Reader;

[TestFixture]
public class ExcelReaderTest
{
    [Test]
    public void GetRow_ReturnsRowAsNewInstanceOfType_WithTheCorrectValues()
    {
        // Assert
        var options = new XLeratorOptions
        {
            FilePath = "./GetRow_ReturnsRowAsNewInstanceOfType_WithTheCorrectValues",
            HeaderLength = 1
        };

        var mapper = IndexedExcelMapper.CreateFrom(typeof(IndexedExcelClass));
        
        var data = new IndexedExcelClass
        {
            Id = 69,
            Name = "Test"
        };

        var creator = ExcelCreator<IndexedExcelClass>.Create(options, mapper);
        using (var reader = creator.CreateExcel())
        {
            reader.Write(data);
        }

        using var testee = ExcelReader<IndexedExcelClass>.Create(options, mapper);

        // Act
        var actual = testee.GetRow(0);
        
        // Assert
        actual.Should().BeEquivalentTo(data);
    }
}