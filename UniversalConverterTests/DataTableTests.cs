using UniversalConverter;

namespace UniversalConverterTests;
public class DataTableTests
{
    [SetUp]
    public void Setup()
    {
    }

    [Test]
    public void SaveAsExcelDocument_Works()
    {
        var list = GetTestList();
        var dataTableFromList = list.Convert().To.DataTable;
        dataTableFromList.Convert().SaveAs("./test.xlsx");
        Assert.Pass();
    }

    [Test]
    public void ConvertToList_Works()
    {
        var list = GetTestList();
        var dataTableFromList = list.Convert().To.DataTable;
        var newList = dataTableFromList.Convert().ToGeneric<TestUser>().List;
        Assert.Pass();
    }
    private static List<TestUser> GetTestList()
    {
        return new List<TestUser>
        {
            new() { Name = "Joe", Age = 10, BirthDay = DateTime.Today },
            new() { Name = "Tom", Age = 10, BirthDay = DateTime.Today.AddDays(1) },
            new() { Name = "Pascal", Age = 10 }
        };
    }

    internal class TestUser
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public DateTime BirthDay { get; set; }
    }
}