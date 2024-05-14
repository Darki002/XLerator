namespace XLerator.Tests;

[SetUpFixture]
public static class TestEnvironment
{
    public static readonly List<string> FilePaths = new List<string>();

    [OneTimeTearDown]
    public static void TearDown()
    {
        foreach (var file in FilePaths)
        {
            try
            {
                if (File.Exists(file))
                {
                    File.Delete(file);
                }
            }
            catch
            {
                Console.WriteLine($"File {file} couldn't be deleted.");
            }
        }
    }
}