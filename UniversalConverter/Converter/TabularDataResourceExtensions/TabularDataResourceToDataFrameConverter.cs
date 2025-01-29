using Microsoft.Data.Analysis;
using Microsoft.DotNet.Interactive.Formatting.TabularData;

namespace UniversalConverter.Converter.TabularDataResourceExtensions;

public class TabularDataResourceToDataFrameConverter
{
    public DataFrame Convert(TabularDataResource tabularDataResource)
    {
        if (tabularDataResource == null)
        {
            throw new ArgumentNullException(nameof(tabularDataResource));
        }

        var dataFrame = new DataFrame();

        foreach (var fieldDescriptor in tabularDataResource.Schema.Fields)
        {
            switch (fieldDescriptor.Type)
            {
                case TableSchemaFieldType.Number:
                    dataFrame.Columns.Add(new DoubleDataFrameColumn(fieldDescriptor.Name, tabularDataResource.Data.Select(d => System.Convert.ToDouble(d.FirstOrDefault(x => x.Key == fieldDescriptor.Name).Value))));
                    break;
                case TableSchemaFieldType.Integer:
                    dataFrame.Columns.Add(new Int64DataFrameColumn(fieldDescriptor.Name, tabularDataResource.Data.Select(d => System.Convert.ToInt64(d.FirstOrDefault(x => x.Key == fieldDescriptor.Name).Value))));
                    break;
                case TableSchemaFieldType.Boolean:
                    dataFrame.Columns.Add(new BooleanDataFrameColumn(fieldDescriptor.Name, tabularDataResource.Data.Select(d => System.Convert.ToBoolean(d.FirstOrDefault(x => x.Key == fieldDescriptor.Name).Value))));
                    break;
                case TableSchemaFieldType.String:
                    dataFrame.Columns.Add(new StringDataFrameColumn(fieldDescriptor.Name, tabularDataResource.Data.Select(d => System.Convert.ToString(d.FirstOrDefault(x => x.Key == fieldDescriptor.Name).Value))));
                    break;
                case TableSchemaFieldType.DateTime:
                    dataFrame.Columns.Add(new DateTimeDataFrameColumn(fieldDescriptor.Name, tabularDataResource.Data.Select(d => System.Convert.ToDateTime(d.FirstOrDefault(x => x.Key == fieldDescriptor.Name).Value))));
                    break;
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }
        return dataFrame;
    }
}