using XLerator.ExcelUtility.Factories;
using XLerator.Mappings;
using XLerator.Tests.TestObjects;

namespace XLerator.Tests.ExcelUtility.Factories;

[TestFixture]
public class XLeratorFactoryTests
{
    [Test]
    public void CreateMapper_ReturnsHeaderExcelMapper_WhenClassHasNoAttribute()
    {
        // Act
        var mapper = XLeratorFactory.CreateMapper(typeof(HeaderedExcelClass));
        
        // Assert
        mapper.Should().BeOfType<HeaderExcelMapper>();
    }
    
    [Test]
    public void CreateMapper_ReturnsIndexedExcelMapper_WhenAttributeIsOnClass()
    {
        // Act
        var mapper = XLeratorFactory.CreateMapper(typeof(IndexedExcelClass));
        
        // Assert
        mapper.Should().BeOfType<IndexedExcelMapper>();
    }
}