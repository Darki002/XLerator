using XLerator.ExcelUtility;

namespace XLerator.Tests.ExcelUtility;

[TestFixture]
public class XLeratorOptionsTest
{
    [Test]
    public void SheetNameOrDefault_ReturnsSheetName()
    {
        // Arrange
        var testee = new XLeratorOptions
        {
            FilePath = "",
            SheetName = "test"
        };
        
        // Act
        var actual = testee.GetSheetNameOrDefault();
        
        // Assert
        actual.Should().Be("test");
    }
    
    [Test]
    public void SheetNameOrDefault_ReturnsDefault_WhenSheetNameIsNull()
    {
        // Arrange
        var testee = new XLeratorOptions
        {
            FilePath = "",
            SheetName = null
        };
        
        // Act
        var actual = testee.GetSheetNameOrDefault();
        
        // Assert
        actual.Should().Be("Sheet1");
    }
}