namespace XLerator.Tests.Mappings;

[TestFixture]
public class ExcelMapperBaseTest
{
    [Test]
    [TestCase(1, "A")]
    [TestCase(3, "C")]
    [TestCase(27, "AA")]
    [TestCase(703, "AAA")]
    public void GetColumnFor_ReturnsExcelColumnString_ForTheCorrespondingPropertyName(int index, string expected)
    {
        // Arrange
        var testee = new ExcelMapperBaseFake();
        testee.AddPropertyIndexMap("Test", index);
        
        // Act
        var excelCol = testee.GetColumnFor("Test");
        
        // Assert
        excelCol.Should().Be(expected);
    }
}