using XLerator.ExcelUtility;
using XLerator.ExcelUtility.ExcelEditing;
using XLerator.Tests.Mappings;
using XLerator.Tests.TestObjects;

namespace XLerator.Tests.ExcelUtility;

[TestFixture]
public class ExcelRowTest
{
    [Test]
    public void CreateHeader_ReturnsNewExcelRow_WithTheCorrectExcelCells()
    {
        // Arrange
        var excelMapper = new ExcelMapperBaseFake();
        excelMapper.AddHeaderMap("Id", "Index");
        excelMapper.AddPropertyIndexMap("Id", 1);
        excelMapper.AddHeaderMap("Name", "Name");
        excelMapper.AddPropertyIndexMap("Name", 2);
        
        var expected1 = new ExcelCell("A", 0, "Index");
        var expected2 = new ExcelCell("B", 0, "Name");
        
        // Act
        var testee = ExcelHeader<HeaderedExcelClass>.CreateFrom(0, excelMapper);
        
        // Assert
        testee.Row.Should().HaveCount(2);
        testee[0].Should().BeEquivalentTo(expected1);
        testee[1].Should().BeEquivalentTo(expected2);
    }
    
    [Test]
    public void CreateFrom_ReturnsNewExcelRow_WithTheCorrectExcelCells()
    {
        // Arrange
        var data = new HeaderedExcelClass
        {
            Id = 1,
            Name = "Test"
        };
        
        var excelMapper = new ExcelMapperBaseFake();
        excelMapper.AddHeaderMap("Id", "Index");
        excelMapper.AddPropertyIndexMap("Id", 1);
        excelMapper.AddHeaderMap("Name", "Name");
        excelMapper.AddPropertyIndexMap("Name", 2);
        
        var expected1 = new ExcelCell("A", 0, data.Id);
        var expected2 = new ExcelCell("B", 0, data.Name);
        
        // Act
        var testee = ExcelData<HeaderedExcelClass>.CreateFrom(data, 0, excelMapper);
        
        // Assert
        testee.Row.Should().HaveCount(2);
        testee[0].Should().BeEquivalentTo(expected1);
        testee[1].Should().BeEquivalentTo(expected2);
    }
}