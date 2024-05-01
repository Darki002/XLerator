namespace XLerator.Tags;

[AttributeUsage(AttributeTargets.Property)]
public class ExcelIndex(int index) : Attribute
{
    public int Index { get; } = index;
}