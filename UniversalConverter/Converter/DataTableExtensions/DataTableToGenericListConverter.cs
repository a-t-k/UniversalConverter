using System.Data;
namespace UniversalConverter.Converter.DataTableExtensions;
public class DataTableToGenericListConverter
{
    public List<T> Convert<T>(DataTable dataTable)
    {
        return (from DataRow row in dataTable.Rows select GetInstance<T>(row)).ToList();
    }

    private static T GetInstance<T>(DataRow dataRow)
    {
        var listType = typeof(T);
        var instance = Activator.CreateInstance<T>();
        foreach (DataColumn column in dataRow.Table.Columns)
        {
            foreach (var propertyInfo in listType.GetProperties())
            {
                if (string.Equals(propertyInfo.Name, column.ColumnName, StringComparison.CurrentCultureIgnoreCase))
                {
                    propertyInfo.SetValue(instance, dataRow[column.ColumnName], null);
                }
            }
        }
        return instance;
    }
}