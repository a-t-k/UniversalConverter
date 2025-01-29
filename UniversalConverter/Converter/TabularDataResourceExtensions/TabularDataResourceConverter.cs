using Microsoft.Data.Analysis;
using Microsoft.DotNet.Interactive.Formatting.TabularData;

namespace UniversalConverter.Converter.TabularDataResourceExtensions;

/// <summary>
///   Converter 
/// </summary>
public class TabularDataResourceConverter(TabularDataResource data)
{
    protected readonly TabularDataResource data = data;

    public (DataFrame DataFrame, object Test) To =>
            (
                DataFrame: new TabularDataResourceToDataFrameConverter().Convert(this.data),
                Test: new()
            );
}