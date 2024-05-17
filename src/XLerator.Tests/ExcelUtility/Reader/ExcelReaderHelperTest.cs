using DocumentFormat.OpenXml.Spreadsheet;
using XLerator.ExcelUtility;
using XLerator.ExcelUtility.Reader;
using XLerator.Tests.Mappings;

namespace XLerator.Tests.ExcelUtility.Reader;

[TestFixture]
public class ExcelReaderHelperTest
{
    [TestFixture]
    private class GetDefaultValueTests
    {
        [Test]
        public void GetDefaultValue_ReturnsNull_WhenTypeIsNullable()
        {
            // Arrange
            var type = typeof(int?);
        
            // Act
            var actual = ExcelReader<Dummy>.GetDefaultValue(type);
        
            // Assert
            actual.Should().BeNull();
        }
    
        [Test]
        public void GetDefaultValue_ReturnsZero_WhenTypeIsInt()
        {
            // Arrange
            var type = typeof(int);
        
            // Act
            var actual = ExcelReader<Dummy>.GetDefaultValue(type);
        
            // Assert
            actual.Should().Be(0);
        }
    
        [Test]
        public void GetDefaultValue_ReturnsEmptyString_WhenTypeIsString()
        {
            // Arrange
            var type = typeof(string);
        
            // Act
            var actual = ExcelReader<Dummy>.GetDefaultValue(type);
        
            // Assert
            actual.Should().Be(string.Empty);
        }
    
        [Test]
        public void GetDefaultValue_ReturnsDateTimeMin_WhenTypeIsDateTime()
        {
            // Arrange
            var type = typeof(DateTime);
        
            // Act
            var actual = ExcelReader<Dummy>.GetDefaultValue(type);
        
            // Assert
            actual.Should().Be(DateTime.MinValue);
        }   
    }
    
    [TestFixture]
    private class GetCellValueTests
    {
        [Test]
        public void GetCellValue_ThrowsArgumentException_WhenNoCellIndexIsFound()
        {
            // Arrange
            var options = new XLeratorOptions
            {
                FilePath = "./GetCellValue_ThrowsArgumentException_WhenNoCellIndexIsFound.xlsx"
            };

            using (var _ = Spreadsheet.Create(options)) { }
            XLeratorTest.FilePaths.Add(options.FilePath);
            
            var testee = ExcelReader<Dummy>.Create(options, new ExcelMapperDummy());
            
            // Act
            var actual = () => testee.GetCellValue(new List<Cell>(), "Test");
            
            // Asset
            actual.Should().Throw<ArgumentException>();
        }
        
        [Test]
        public void GetCellValue_ReturnsCorrectCell()
        {
            // Arrange
            var options = new XLeratorOptions
            {
                FilePath = "./GetCellValue_ReturnsCorrectCell.xlsx"
            };

            using (var _ = Spreadsheet.Create(options)) { }
            XLeratorTest.FilePaths.Add(options.FilePath);

            var mapper = new ExcelMapperBaseFake();
            mapper.AddPropertyIndexMap("Test", 2);
            
            var testee = ExcelReader<Dummy>.Create(options, mapper);

            var cell = new Cell
            {
                CellValue = new CellValue("Test")
            };

            var cells = new List<Cell>
            {
                new Cell { CellValue = new CellValue("fdsfasdfa")},
                cell
            };
            
            // Act
            var actual = testee.GetCellValue(cells, "Test");
            
            // Asset
            actual.Should().Be("Test");
        }
    }
    
    [TestFixture]
    private class GetValueOrDefaultTests
    {
        [Test]
        public void GetValueOrDefault_ReturnsDefaultValue_WhenStringIsNull()
        {
            // Act
            var actual = ExcelReader<Dummy>.GetValueOrDefault(typeof(int), null);
            
            // Asset
            actual.Should().Be(0);
        }
        
        [Test]
        public void GetValueOrDefault_ReturnsValueFromString()
        {
            // Act
            var actual = ExcelReader<Dummy>.GetValueOrDefault(typeof(int), "69");
            
            // Asset
            actual.Should().BeOfType<int>();
            actual.Should().Be(69);
        }
    }
}