using XLerator.ExcelUtility;

namespace XLerator.Tests.ExcelUtility;

[TestFixture]
public class XLeratorOptionsTest
{
    [Test]
    public void GetFilePath_ThrowsIfFilePathIsNull()
    {
        // Arrange
        var testee = new XLeratorOptions
        {
            FilePath = null
        };
        
        // Act
        var actual = () => testee.GetFilePath();
        
        // Assert
        actual.Should().Throw<ArgumentNullException>();
    }
    
    [Test]
    public void GetFilePath_ReturnsTheFilePath()
    {
        // Arrange
        var testee = new XLeratorOptions
        {
            FilePath = "test"
        };
        
        // Act
        var actual = testee.GetFilePath();
        
        // Assert
        actual.Should().Be("test");
    }

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