using XLerator.Tests.Mappings;

namespace XLerator.Tests.ExcelMappings;

[TestFixture]
public class ExcelMapperBaseTest
{
    [Test]
    [TestCase(1, "A")]
    [TestCase(3, "C")]
    [TestCase(27, "AA")]
    [TestCase(703, "AAA")]
    public void GetColumnFor_ReturnsExcelColumnString_ForTheCorrespondingPropertyName(int index, string expected)
    {
        // Arrange
        var testee = new ExcelMapperBaseFake();
        testee.AddPropertyIndexMap("Test", index);
        
        // Act
        var excelCol = testee.GetColumnFor("Test");
        
        // Assert
        excelCol.Should().Be(expected);
    }

    [Test]
    public void GetHeaderNameFor_ReturnsHeaderName_WhenHeaderIsForPropertyDefined()
    {
        // Arrange
        var testee = new ExcelMapperBaseFake();
        testee.AddHeaderMap("Test", "Test Header");
        
        // Act
        var excelCol = testee.GetHeaderFor("Test");
        
        // Assert
        excelCol.Should().Be("Test Header");
    }
    
    [Test]
    public void GetHeaderNameFor_ReturnsPropertyName_WhenNoHeaderIsDefined()
    {
        // Arrange
        var testee = new ExcelMapperBaseFake();
        
        // Act
        var excelCol = testee.GetHeaderFor("Test");
        
        // Assert
        excelCol.Should().Be("Test");
    }
}