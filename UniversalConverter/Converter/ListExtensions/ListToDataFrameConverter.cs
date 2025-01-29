using Microsoft.Data.Analysis;

namespace UniversalConverter.Converter.ListExtensions;
public class ListToDataFrameConverter
{
    public DataFrame Convert<T>(List<T> list)
    {
        var properties = typeof(T).GetProperties();
        var values = list.Select(t => properties.Select(p => p.GetValue(t)).ToList());
        var columnInfos = properties.Select(p => (p.Name, p.PropertyType)).ToList();
        var dataFrame = DataFrame.LoadFrom(values, columnInfos);
        return dataFrame;
    }
}