using Microsoft.DotNet.Interactive.Formatting.TabularData;
using UniversalConverter;
namespace UniversalConverterTests;
public class TabularDataResourceToDataFrameTest
{
    [SetUp]
    public void Setup()
    {
    }

    [Test]
    public void Convert_Works()
    {
        var tableSchema = new TableSchema();
        var data = Array.Empty<IEnumerable<KeyValuePair<string, object>>>();
        var tabularData = new TabularDataResource(tableSchema, data);
        var dataFrame = tabularData.Convert().To.DataFrame;
        Assert.Pass();
    }
}