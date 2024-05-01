namespace XLerator;

public class XLeratorFactory(string filePath)
{
    public ExcelReader CreateReader<T>() where T : class
    {
        return ExcelReader.Create<T>(filePath);
    }

    public ExcelWriter CreateWriter<T>() where T : class
    {
        return ExcelWriter.Create<T>(filePath);
    }
}