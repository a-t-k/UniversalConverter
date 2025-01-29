using Microsoft.DotNet.Interactive.Formatting.TabularData;
using System.Data;
using UniversalConverter.Converter.DataTableExtensions;
using UniversalConverter.Converter.ListExtensions;
using UniversalConverter.Converter.TabularDataResourceExtensions;

namespace UniversalConverter;

public static class ExtensionConverterWrapper
{
    public static TabularDataResourceConverter Convert(this TabularDataResource data) => new(data);
    public static ListConverter<T> Convert<T>(this List<T> data) => new(data);
    public static DataTableConverter Convert(this DataTable data) => new(data);
}