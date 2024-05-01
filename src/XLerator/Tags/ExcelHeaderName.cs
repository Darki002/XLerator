namespace XLerator.Tags;

[AttributeUsage(AttributeTargets.Property)]
public class ExcelHeaderName(string name) : Attribute
{
    public string Name { get; } = name;
}