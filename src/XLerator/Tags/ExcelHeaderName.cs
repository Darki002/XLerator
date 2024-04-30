namespace XLerator.Tags;

[AttributeUsage(AttributeTargets.Property)]
public abstract class ExcelHeaderName(string name) : Attribute
{
    public string Name { get; } = name;
}