namespace XLerator.Mappings;

internal abstract class ExcelMapperBase
{
    internal readonly Dictionary<string, string> HeaderMap = new Dictionary<string, string>();
    internal readonly Dictionary<string, int> PropertyIndexMap = new Dictionary<string, int>();

    public abstract string? GetHeaderFor(string propertyName);

    public string? GetColumnFor(string propertyName)
    {
        if (PropertyIndexMap.TryGetValue(propertyName, out var columnNumber))
        {
            return IntToColumnString(columnNumber);
        }

        return null;
    }

    public int? GetColumnIndexFor(string propertyName) =>
        PropertyIndexMap.TryGetValue(propertyName, out var index) ? index : null;

    private static string IntToColumnString(int columnNumber)
    {
        if (columnNumber <= 0)
        {
            throw new ArgumentException("Column Index must be greater then Zero.");
        }

        var columnName = string.Empty;
        while (columnNumber > 0)
        {
            var remainder = (columnNumber - 1) % 26;
            columnName = (char)(remainder + 'A') + columnName;
            columnNumber = (columnNumber - 1) / 26;
        }

        return columnName;
    }

    // TODO: allow to add old mappings, in case the class used to be different, so that the old data still can be read.
}