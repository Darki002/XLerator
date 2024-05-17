using XLerator.ExcelUtility.Reader;

namespace XLerator.Tests.ExcelUtility.Reader;

[TestFixture]
public class HelperTests
{
    [Test]
    public void GetDefaultValue_ReturnsNull_WhenTypeIsNullable()
    {
        // Arrange
        var type = typeof(int?);
        
        // Act
        var actual = Helper.GetDefaultValue(type);
        
        // Assert
        actual.Should().BeNull();
    }
    
    [Test]
    public void GetDefaultValue_ReturnsZero_WhenTypeIsInt()
    {
        // Arrange
        var type = typeof(int);
        
        // Act
        var actual = Helper.GetDefaultValue(type);
        
        // Assert
        actual.Should().Be(0);
    }
    
    [Test]
    public void GetDefaultValue_ReturnsEmptyString_WhenTypeIsString()
    {
        // Arrange
        var type = typeof(string);
        
        // Act
        var actual = Helper.GetDefaultValue(type);
        
        // Assert
        actual.Should().Be(string.Empty);
    }
    
    [Test]
    public void GetDefaultValue_ReturnsDateTimeMin_WhenTypeIsDateTime()
    {
        // Arrange
        var type = typeof(DateTime);
        
        // Act
        var actual = Helper.GetDefaultValue(type);
        
        // Assert
        actual.Should().Be(DateTime.MinValue);
    }
}