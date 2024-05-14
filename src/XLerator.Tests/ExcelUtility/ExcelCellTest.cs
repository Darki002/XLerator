using DocumentFormat.OpenXml.Spreadsheet;
using XLerator.ExcelUtility;

namespace XLerator.Tests.ExcelUtility;

[TestFixture]
public class ExcelCellTest
{
    [Test]
    public void ToCell_WithNonNullData_ShouldSetCellValueCorrectly()
    {
        // Arrange
        var excelCell = new ExcelCell("A", 1, "Hello");

        // Act
        var cell = excelCell.ToCell();

        // Assert
        cell.CellValue!.Text.Should().Be("Hello");
        cell.DataType.Should().Be(CellValues.String);
        cell.CellReference!.Value.Should().Be("A1");
    }

    [Test]
    public void ToCell_WithNumericData_ShouldSetDataTypeToNumber()
    {
        // Arrange
        var excelCell = new ExcelCell("B", 2, 123);

        // Act
        var cell = excelCell.ToCell();

        // Assert
        cell.DataType.Should().Be(CellValues.Number);
    }

    [Test]
    public void ToCell_WithBooleanData_ShouldSetDataTypeToBoolean()
    {
        // Arrange
        var excelCell = new ExcelCell("C", 3, true);

        // Act
        var cell = excelCell.ToCell();

        // Assert
        cell.DataType.Should().Be(CellValues.Boolean);
    }

    [Test]
    public void ToCell_WithDateTimeData_ShouldSetDataTypeToDate()
    {
        // Arrange
        var dateTime = new DateTime(2020, 1, 1);
        var excelCell = new ExcelCell("D", 4, dateTime);

        // Act
        var cell = excelCell.ToCell();

        // Assert
        cell.DataType.Should().Be(CellValues.Date);
    }
}