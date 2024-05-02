namespace XLerator.Attributes;

/// <summary>
/// Defines the Text for this Property that will be used for the Header, instead of the Property Name.
/// </summary>
/// <param name="name">The Text for the Header</param>
[AttributeUsage(AttributeTargets.Property)]
public class ExcelHeaderName(string name) : Attribute
{
    public string Name { get; } = name;
}