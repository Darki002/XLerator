namespace XLerator.Attributes;

/// <summary>
/// Defines the Column Index for the Property. 
/// </summary>
/// <param name="index">The 1 based index of the column. For Example B = 2</param>
[AttributeUsage(AttributeTargets.Property)]
public class ColumnIndex(int index) : Attribute
{
    public int Index { get; } = index;
}