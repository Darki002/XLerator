namespace XLerator.ExcelUtility.Factories;

public partial class XLeratorFactory<T>
{
    private readonly XLeratorOptions options;
    
    private XLeratorFactory(XLeratorOptions options)
    {
        this.options = options;
    }
}