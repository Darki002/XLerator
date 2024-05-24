using XLerator.ExcelUtility;
using XLerator.ExcelUtility.ExcelReading.Iterator;
using XLerator.Mappings;
using XLerator.Tests.TestObjects;

namespace XLerator.Tests.ExcelUtility.ExcelReadingTest.Iterator;

[TestFixture]
[NonParallelizable]
public class IteratorTest
{
    [Test]
    public void Read_ReturnsTrue_WhenThereAreMoreRowsLeft()
    {
        // Arrange
        var iterator = CreateIterator();
        
        // Act
        var actual = iterator.Read();
        
        // Asser
        actual.Should().BeTrue();
    }
    
    [Test]
    public void Read_ReturnsFalse_WhenThereAreNoMoreRowsLeft()
    {
        // Arrange
        var iterator = CreateIterator();
        
        // Act
        for (var i = 0; i < 3; i++)
        {
            iterator.Read();
        }
        var actual = iterator.Read();
        
        // Asser
        actual.Should().BeFalse(); 
    }

    [Test]
    public void GetCurrentRow_ReturnsCurrentSelectedRow_AsNewInstance()
    {
        // Arrange
        var expected = new IndexedExcelClass
        {
            Id = 1,
            Name = "Test"
        };
        
        var iterator = CreateIterator();
        iterator.Read();
        
        // Act
        var actual = iterator.GetCurrentRow();
        
        // Assert
        actual.Should().BeEquivalentTo(expected);
    }

    [Test]
    public void SkipRows_SkipsTheGivenAmountOfRowsInSpreadSheet()
    {
        // Arrange
        var expected = new IndexedExcelClass
        {
            Id = 3,
            Name = "Test3"
        };
        
        var iterator = CreateIterator();
        
        // Act
        iterator.SkipRows(2);
        
        // Assert
        var actual = iterator.GetCurrentRow();
        actual.Should().BeEquivalentTo(expected);
    }
    
    [Test]
    public void SkipRows_ThrowsArgumentException_WhenAmountIsEqualOrLessThenZero()
    {
        // Arrange
        var iterator = CreateIterator();
        
        // Act
        var actual = () => iterator.SkipRows(-2);
        
        // Assert
        actual.Should().Throw<ArgumentException>();
    }

    private static ExcelIterator<IndexedExcelClass> CreateIterator()
    {
        var options = new XLeratorOptions
        {
            FilePath = "./ExcelUtility/ExcelReadingTest/Iterator/test.xlsx",
        };

        var mapper = IndexedExcelMapper.CreateFrom(typeof(IndexedExcelClass));

        return ExcelIterator<IndexedExcelClass>.Create(options, mapper);
    }
}