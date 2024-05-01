namespace XLerator.Attributes;

[AttributeUsage(AttributeTargets.Property)]
public class ExcelHeaderName(string name) : Attribute
{
    public string Name { get; } = name;
}