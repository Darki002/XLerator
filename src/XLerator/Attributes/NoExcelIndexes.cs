namespace XLerator.Attributes;

/// <summary>
///     Defines a class to not use specific column indexes for the spreadsheet. Instead, it will use the order of the public
///     properties as the column index.
/// </summary>
[AttributeUsage(AttributeTargets.Class)]
public class NoExcelIndexes : Attribute;