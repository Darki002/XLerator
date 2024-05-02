namespace XLerator.Attributes;

/// <summary>
/// Defines the Index for the Property. 
/// </summary>
/// <param name="index">The index</param>
[AttributeUsage(AttributeTargets.Property)]
public class ExcelIndex(int index) : Attribute
{
    public int Index { get; } = index;
}