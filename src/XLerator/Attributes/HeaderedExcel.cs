namespace XLerator.Attributes;

/// <summary>
///     Defines a class to use the Header to determine which Property belongs to which Column in the Spreadsheet.
///     Use the <see cref="HeaderName"/> Attribute to set the header name.
/// </summary>
[AttributeUsage(AttributeTargets.Class)]
public class HeaderedExcel : Attribute;