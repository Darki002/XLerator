using XLerator.Mappings;
using XLerator.Tests.TestObjects;

namespace XLerator.Tests.ExcelMappings;

[TestFixture]
public class IndexedExcelMapperTest
{
    [Test]
    public void CreateFrom_ReturnsNewInstance_WithCorrectPropertyIndexMap()
    {
        // Arrange
        var expected = new Dictionary<string, int>
        {
            { nameof(IndexedExcelClass.Id), 1 },
            { nameof(IndexedExcelClass.Name), 2 }
        };
        
        // Act
        var testee = IndexedExcelMapper.CreateFrom(typeof(IndexedExcelClass));
        
        // Assert
        testee.PropertyIndexMap.Should().BeEquivalentTo(expected);
    }
    
    [Test]
    public void CreateFrom_ReturnsNewInstance_WithCorrectHeaderMap()
    {
        // Arrange
        var expected = new Dictionary<string, string>
        {
            { nameof(IndexedExcelClass.Id), "Index" }
        };
        
        // Act
        var testee = IndexedExcelMapper.CreateFrom(typeof(IndexedExcelClass));
        
        // Assert
        testee.HeaderMap.Should().BeEquivalentTo(expected);
    }
}