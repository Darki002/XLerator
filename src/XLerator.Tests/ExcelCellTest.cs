using DocumentFormat.OpenXml.Spreadsheet;
using XLerator.ExcelUtility;

namespace XLerator.Tests;

[TestFixture]
public class ExcelCellTest
{
    [Test]
    public void ToCell_WithNonNullData_ShouldSetCellValueCorrectly()
    {
        // Arrange
        var excelCell = new ExcelCell<string>("A", 1, "Hello");

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
        var excelCell = new ExcelCell<int>("B", 2, 123);

        // Act
        var cell = excelCell.ToCell();

        // Assert
        cell.DataType.Should().Be(CellValues.Number);
    }

    [Test]
    public void ToCell_WithBooleanData_ShouldSetDataTypeToBoolean()
    {
        // Arrange
        var excelCell = new ExcelCell<bool>("C", 3, true);

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
        var excelCell = new ExcelCell<DateTime>("D", 4, dateTime);

        // Act
        var cell = excelCell.ToCell();

        // Assert
        cell.DataType.Should().Be(CellValues.Date);
    }

    [Test]
    public void ToCell_WithDataNull_ShouldThrowInvalidOperationException()
    {
        // Arrange
        var excelCell = new ExcelCell<string>("E", 5);

        // Act
        Action act = () => excelCell.ToCell();

        // Assert
        act.Should().Throw<InvalidOperationException>()
            .WithMessage("Not Data to convert to CellValue");
    }
}