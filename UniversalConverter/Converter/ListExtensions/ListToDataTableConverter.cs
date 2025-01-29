using System.Data;
namespace UniversalConverter.Converter.ListExtensions;
public class ListToDataTableConverter
{
    public DataTable Convert<T>(List<T> list)
    {
        var dt = new DataTable();
        foreach (var info in typeof(T).GetProperties())
        {
            dt.Columns.Add(new DataColumn(info.Name, GetNullableType(info.PropertyType)));
        }
        foreach (var t in list)
        {
            var row = dt.NewRow();
            foreach (var info in typeof(T).GetProperties())
            {
                row[info.Name] = IsNullableType(info.PropertyType)
                    ? info.GetValue(t, null) ?? DBNull.Value
                    : info.GetValue(t, null);
            }

            dt.Rows.Add(row);
        }
        return dt;
    }
    private static Type GetNullableType(Type t)
    {
        var returnType = t;
        if (t.IsGenericType && t.GetGenericTypeDefinition() == typeof(Nullable<>))
        {
            returnType = Nullable.GetUnderlyingType(t);
        }

        return returnType;
    }
    private static bool IsNullableType(Type type)
    {
        return (type == typeof(string)
                || type.IsArray
                || (type.IsGenericType && type.GetGenericTypeDefinition() == typeof(Nullable<>)));
    }
}