using Ninjasoft.Excel;

namespace Ninjasoft.Excel.IntegrationTests;

public class Tests
{
    [SetUp]
    public void Setup()
    {
    }

    [Test]
    public void Test1()
    {
        ExcelFile excelFile = new(@"/home/t/Downloads/Demo.xlsx");
        excelFile.Read();
        //Assert.That(result, Is.True);
    }
}
