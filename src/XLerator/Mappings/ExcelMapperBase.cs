namespace XLerator.Mappings;

internal abstract class ExcelMapperBase
{ 
     internal readonly Dictionary<string, int> PropertyIndexMap = new Dictionary<string, int>();

     internal readonly Dictionary<string, string> HeaderMap = new Dictionary<string, string>();

     public abstract (string Name, int Index)? GetHeaderFor(string propertyName);
     
     public string? GetColumnFor(string propertyName)
     {
          if (PropertyIndexMap.TryGetValue(propertyName, out var columnNumber))
          {
               return IntToColumnString(columnNumber);
          }

          return null;
     }

     private static string IntToColumnString(int columnNumber)
     {
          var columnName = string.Empty;
          while (columnNumber > 0)
          {
               var remainder = (columnNumber - 1) % 26;
               columnName = (char)(remainder + 'A') + columnName;
               columnNumber = (columnNumber - 1) / 26;
          }
          return columnName;
     }
}