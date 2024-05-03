using XLerator.ExcelUtility;
using XLerator.ExcelUtility.Creator;

namespace XLerator.Tests.ExcelUtility.Creator;

[TestFixture]
public class ExcelCreatorTest
{
    [Test]
    public void CreateExcel_ReturnsNewIExcelEditor()
    {
        // Arrange
        const string filePath = "./CreateExcel_ReturnsNewIExcelEditor.xlsx";
        var options = new XLeratorOptions
        {
            FilePath = filePath
        };

        var testee = ExcelCreator<Dummy>.Create(options, new ExcelMapperDummy());
        
        // Act
        var excelEditor = testee.CreateExcel(false);
        
        // Assert
        excelEditor.Should().NotBeNull();
        excelEditor.Dispose();
        
        // Clean Up
        if (File.Exists(filePath))
        {
            File.Delete(filePath);
        }
    }
    
    [Test]
    public void CreateExcel_CreatesANewExcelFile()
    {
        // Arrange
        const string filePath = "./CreateExcel_CreatesANewExcelFile.xlsx";
        var options = new XLeratorOptions
        {
            FilePath = filePath
        };

        var testee = ExcelCreator<Dummy>.Create(options, new ExcelMapperDummy());
        
        // Act
        var editor = testee.CreateExcel(false);
        editor.Dispose();
        
        // Assert
        var fileExist = File.Exists(filePath);
        fileExist.Should().BeTrue();
        
        // Clean Up
        if (fileExist)
        {
            File.Delete(filePath);
        }
        fileExist = File.Exists(filePath);
        fileExist.Should().BeFalse();
    }
}