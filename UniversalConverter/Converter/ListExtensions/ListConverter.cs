using Microsoft.Data.Analysis;
using System.Data;

namespace UniversalConverter.Converter.ListExtensions;
public class ListConverter<T>(List<T> data)
{
    protected readonly List<T> data = data;

    public (DataTable DataTable, DataFrame DataFrame) To =>
    (
        DataTable: new ListToDataTableConverter().Convert(this.data),
        DataFrame: new ListToDataFrameConverter().Convert<T>(this.data)
    );
}