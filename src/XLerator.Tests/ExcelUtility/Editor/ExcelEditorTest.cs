using XLerator.ExcelUtility;
using XLerator.ExcelUtility.Creator;
using XLerator.ExcelUtility.Editor;

namespace XLerator.Tests.ExcelUtility.Editor;

[TestFixture]
public class ExcelEditorTest
{
    [Test]
    public void Create_InitializedCorrectly()
    {
        // Arrange
        const string filePath = "./CreateExcel_ReturnsNewIExcelEditor.xlsx";
        var options = new XLeratorOptions
        {
            FilePath = filePath
        };
        var creator = ExcelCreator<Dummy>.Create(options, new ExcelMapperDummy());
        
        // Act
        var testee = creator.CreateExcel(false) as ExcelEditor<Dummy>;
        
        // Assert
        testee.Should().NotBeNull();
        testee!.Spreadsheet.Should().NotBeNull();
        testee.SheetData.Should().NotBeNull();
        
        // Clean Up
        testee.Dispose();
        
        // Act
        testee = ExcelEditor<Dummy>.Create(options, new ExcelMapperDummy());
        
        // Asser
        testee.Spreadsheet.Should().NotBeNull();
        testee.SheetData.Should().NotBeNull();
        
        // Clean Up
        testee.Dispose();
        if (File.Exists(filePath))
        {
            File.Delete(filePath);
        }
    }
}