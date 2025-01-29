using UniversalConverter;

namespace UniversalConverterTests;
public class ListToDataTableTests
{
    [SetUp]
    public void Setup()
    {
    }

    [Test]
    public void ListToDataTableConvert_Works()
    {
        var list = new List<TestUser>
        {
            new() { Name = "Name", Age = 10}
        };

        var dataTableFromList = list.Convert().To.DataTable;
        Assert.Pass();
    }

    [Test]
    public void ListToDataFrameConvert_Works()
    {
        var list = new List<TestUser>
        {
            new() { Name = "Name", Age = 10, BirthDate = DateTime.Today}
        };

        var dataFrameFromList = list.Convert().To.DataFrame;
        Assert.Pass();
    }

    internal class TestUser
    {
        public string Name { get; set; }
        public int Age { get; set; }
        public DateTime BirthDate { get; set; }
    }
}