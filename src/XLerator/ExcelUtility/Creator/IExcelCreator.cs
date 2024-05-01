namespace XLerator.ExcelUtility.Creator;

public interface IExcelCreator<in T> : IDisposable where T : class
{
    void CreateExcel(IEnumerable<T> rows);
}