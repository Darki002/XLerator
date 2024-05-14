using XLerator.ExcelUtility;
using XLerator.ExcelUtility.Factories;
using XLerator.Tests.TestObjects;

namespace XLerator.Tests;

[TestFixture]
public class Tester
{
    [Test]
    public void Create()
    {
        if (true)
        {
            return;
        }
        
        var options = new XLeratorOptions()
        {
            FilePath = "./Test-file.xlsx",
            SheetName = "Testy"
        };
        
        var factory = XLeratorFactory<HeaderedExcelClass>.CreateFactory(options);
        var creator = factory.CreateExcelCreator();

        using var editor = creator.CreateExcel(true);

        var data = new HeaderedExcelClass
        {
            Id = 42,
            Name = "The answer to everything!"
        };

        var multyData = new List<HeaderedExcelClass>
        {
            new HeaderedExcelClass
            {
                Id = 1,
                Name = "Ka"
            },
            new HeaderedExcelClass
            {
                Id = 69,
                Name = "Hehe"
            }
        };
        
        editor.Write(data);
        editor.WriteMany(multyData);
    }
}