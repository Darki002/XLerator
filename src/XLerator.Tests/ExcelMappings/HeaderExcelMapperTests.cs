using XLerator.ExcelMappings;
using XLerator.Tests.TestObjects;

namespace XLerator.Tests.ExcelMappings;

[TestFixture]
public class HeaderExcelMapperTests
{
    [Test]
    public void CreateFrom_ReturnsNewInstance_WithCorrectPropertyIndexMap()
    {
        // Arrange
        var expected = new Dictionary<string, int>
        {
            { nameof(HeaderedExcelClass.Id), 1 },
            { nameof(HeaderedExcelClass.Name), 2 }
        };
        
        // Act
        var testee = HeaderExcelMapper.CreateFrom(typeof(HeaderedExcelClass));
        
        // Assert
        testee.PropertyIndexMap.Should().BeEquivalentTo(expected);
    }
    
    [Test]
    public void CreateFrom_ReturnsNewInstance_WithCorrectHeaderMap()
    {
        // Arrange
        var expected = new Dictionary<string, string>
        {
            { nameof(HeaderedExcelClass.Id), "Index" },
            { nameof(HeaderedExcelClass.Name), "Name" }
        };
        
        // Act
        var testee = HeaderExcelMapper.CreateFrom(typeof(HeaderedExcelClass));
        
        // Assert
        testee.HeaderMap.Should().BeEquivalentTo(expected);
    }
}