using System.Collections;

namespace XLerator.ExcelUtility;

internal abstract class ExcelRow : IEnumerable<ExcelCell>
{
    internal readonly List<ExcelCell> Row;
    
    internal ExcelRow()
    {
        Row = new List<ExcelCell>();
    }

    internal ExcelCell this[int index] => Row[index];

    public IEnumerator<ExcelCell> GetEnumerator()
    {
        return Row.GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }
}