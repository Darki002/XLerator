namespace XLerator.ExcelUtility;

public class XLeratorOptions(string filePath, string? workbookName)
{
    public readonly string FilePath = filePath;

    public readonly string? WorkbookName = workbookName;
}