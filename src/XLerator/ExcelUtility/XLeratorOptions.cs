namespace XLerator.ExcelUtility;

/// <summary>
/// Options for the Excel file.
/// </summary>
public class XLeratorOptions
{
    /// <summary>
    /// Filepath to the Excel file or where the Excel file should be saved.
    /// </summary>
    public required string FilePath { get; set; }

    /// <summary>
    /// The name for the Sheet which will be read or created.
    /// </summary>
    public string? SheetName { get; set; } = null;
    
    /// <summary>
    /// How many rows are considered as a header and have to be ignored. Default is Zero
    /// </summary>
    public int HeaderLength { get; set; } = 0;

    /// <summary>
    /// The off set for where the first row is located in the Sheet.
    /// </summary>
    public uint RowOffSet { get; set; } = 0;

    /// <summary>
    /// The off set for where the first column is located in the Sheet.
    /// </summary>
    public uint ColumnOffSet { get; set; } = 0;

    internal string GetSheetNameOrDefault() => SheetName ?? "Sheet1";
}