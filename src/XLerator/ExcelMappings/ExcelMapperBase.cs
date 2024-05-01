namespace XLerator.ExcelMappings;

internal abstract class ExcelMapperBase
{ 
     public abstract string GetColumnFor(string propertyName);

     public abstract string GetHeaderNameFor(string propertyName);

     protected static string IntoToColumnString(int columnNumber)
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