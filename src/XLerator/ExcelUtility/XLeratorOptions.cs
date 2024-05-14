namespace XLerator.ExcelUtility;

/// <summary>
/// Options for the Excel file.
/// </summary>
public class XLeratorOptions
{
    /// <summary>
    /// Filepath to the Excel file or where the Excel file should be saved.
    /// </summary>
    public required string FilePath { get; set; } = null;

    /// <summary>
    /// The name for the Sheet which will be read or created.
    /// </summary>
    public string? SheetName { get; set; } = null;

    /// <summary>
    /// The off set for where the first row is located in the Sheet.
    /// </summary>
    public uint rowOffSet { get; set; } = 0;

    /// <summary>
    /// The off set for where the first column is located in the Sheet.
    /// </summary>
    public uint columnOffSet { get; set; } = 0;

    internal string GetFilePath()
    {
        if (FilePath is null)
        {
            throw new ArgumentNullException(nameof(FilePath), "File Path is requered.");
        }

        return FilePath;
    }

    internal string GetSheetNameOrDefault() => SheetName ?? "Sheet1";
}