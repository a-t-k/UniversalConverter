using Microsoft.Data.Analysis;
using System.Data;
using UniversalConverter.Converter.ExcelExtensions;
namespace UniversalConverter.Converter.DataTableExtensions;
public class DataTableConverter(DataTable data)
{
    protected readonly DataTable data = data;

    public (DataFrame DataFrame, object Test) To =>
    (
        new DataTableToDataFrameConverter().Convert(this.data),
        Test: new()
    );

    public (List<T> List, object Test) ToGeneric<T>() =>
    (
        new DataTableToGenericListConverter().Convert<T>(this.data),
        Test: new()
    );

    public (bool ExcelDocument, object Test) SaveAs(string fileName) =>
    (
        new ExcelDocumentCreator().Save(this.data, fileName),
        Test: new()
    );


}