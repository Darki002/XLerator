namespace XLerator.ExcelMappings;

public abstract class ExcelMapperBase
{ 
     protected readonly Dictionary<string, int> PropertyIndexMap = new Dictionary<string, int>();

     protected readonly Dictionary<string, string> HeaderMap = new Dictionary<string, string>();
     
     public string GetColumnFor(string propertyName)
     {
          var columnNumber = PropertyIndexMap[propertyName];
          return IntToColumnString(columnNumber);
     }

     public string GetHeaderNameFor(string propertyName)
     {
          return HeaderMap.GetValueOrDefault(propertyName, propertyName);
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