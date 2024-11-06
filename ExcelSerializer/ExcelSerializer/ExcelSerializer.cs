using ExcelDataReader;
using System.Collections;
using System.Data;
using System.Reflection;
using System.Text;

namespace ExcelSerializer;

public class ExcelSerializer
{
    public static void SerializeToExcel<T>(T data, string filePath)
    {
        SerializeToExcel(data, filePath, ExcelSerializerOptions.Default);
    }

    public static T DeserializeFromExcel<T>(string filePath)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        return DeserializeFromExcel<T>(filePath, ExcelSerializerOptions.Default);
    }

    public static void SerializeToExcel<T>(T data, string filePath, ExcelSerializerOptions options)
    {
        var serializer = new ExcelSerializer();
        serializer.Serialize(data, filePath, options);
    }

    public static T DeserializeFromExcel<T>(string filePath, ExcelSerializerOptions options)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        var serializer = new ExcelSerializer();
        return serializer.Deserialize<T>(filePath, options);
    }

    public void Serialize<T>(T data, string filePath, ExcelSerializerOptions options = null)
    {
        if (data == null) throw new ArgumentException("Data cannot be null.");
        options ??= ExcelSerializerOptions.Default;

        // Check for file extension to determine the format
        var fileExtension = Path.GetExtension(filePath)?.ToLowerInvariant();
        if (fileExtension == ".csv")
        {
            using var writer = new StreamWriter(filePath);
            if (data is IEnumerable dataCollection)
            {
                SerializeCollection(writer, dataCollection, options);
            }
            else
            {
                SerializeSingleObject(writer, data, options);
            }
        }
        else if (fileExtension == ".xlsx")
        {
            throw new NotSupportedException("Serialization to .xlsx is not supported. Please use CSV format.");
        }
        else
        {
            throw new NotSupportedException("Unsupported file format.");
        }
    }

    public T Deserialize<T>(string filePath, ExcelSerializerOptions options = null)
    {
        options ??= ExcelSerializerOptions.Default;

        if (options.DataStartRowIndex < 0)
            throw new ArgumentException(nameof(options.DataStartRowIndex));

        // Check for file extension to determine the format
        var fileExtension = Path.GetExtension(filePath)?.ToLowerInvariant();
        if (fileExtension == ".csv")
        {
            return DeserializeFromCsv<T>(filePath, options);
        }
        else if (fileExtension == ".xlsx")
        {
            return DeserializeFromXlsx<T>(filePath, options);
        }
        else
        {
            throw new NotSupportedException("Unsupported file format.");
        }
    }

    private void SerializeCollection(StreamWriter writer, IEnumerable dataCollection, ExcelSerializerOptions options)
    {
        var itemType = dataCollection.GetType().GetGenericArguments().FirstOrDefault();
        var properties = itemType?.GetProperties();

        if (properties != null)
        {
            writer.WriteLine(string.Join(",", properties.Select(p => p.Name))); // Write header
        }

        foreach (var item in dataCollection)
        {
            var values = properties.Select(p => FormatValue(p.GetValue(item)));
            writer.WriteLine(string.Join(",", values));
        }
    }

    private void SerializeSingleObject<T>(StreamWriter writer, T data, ExcelSerializerOptions options)
    {
        var properties = typeof(T).GetProperties();

        writer.WriteLine(string.Join(",", properties.Select(p => p.Name))); // Write header

        var values = properties.Select(p => FormatValue(p.GetValue(data)));
        writer.WriteLine(string.Join(",", values));
    }

    private string FormatValue(object value)
    {
        return value?.ToString().Replace(",", ";"); // Replace commas in values to avoid CSV issues
    }

    private T DeserializeFromCsv<T>(string filePath, ExcelSerializerOptions options)
    {
        var lines = File.ReadAllLines(filePath);
        var properties = typeof(T).GetProperties().ToDictionary(p => p.Name, p => p);

        if (lines.Length < 2) throw new Exception("Insufficient data in the CSV file.");

        var headers = lines[0].Split(',');
        var dataLine = lines[1].Split(',');

        var obj = Activator.CreateInstance<T>();

        for (int j = 0; j < headers.Length; j++)
        {
            string header = headers[j].Trim();
            if (properties.TryGetValue(header, out var property))
            {
                if (j < dataLine.Length)
                {
                    var cellValue = dataLine[j].Trim();
                    if (string.IsNullOrWhiteSpace(cellValue) && options.TreatEmptyAsNull)
                    {
                        property.SetValue(obj, null);
                    }
                    else
                    {
                        SetPropertyValue(obj, property, cellValue, options);
                    }
                }
            }
        }

        return obj;
    }

    private T DeserializeFromXlsx<T>(string filePath, ExcelSerializerOptions options)
    {
        using var stream = File.Open(filePath, FileMode.Open, FileAccess.Read);
        using var reader = ExcelReaderFactory.CreateReader(stream);

        var dataSet = reader.AsDataSet();
        var table = dataSet.Tables[0]; // Assume we are working with the first sheet

        return typeof(IEnumerable).IsAssignableFrom(typeof(T))
            ? (T)DeserializeCollection<T>(table, options)
            : DeserializeSingleObject<T>(table, options);
    }

    private object DeserializeCollection<T>(DataTable table, ExcelSerializerOptions options)
    {
        var itemType = typeof(T).GetGenericArguments().FirstOrDefault();
        var collectionType = typeof(List<>).MakeGenericType(itemType);
        var collection = (IList)Activator.CreateInstance(collectionType);

        var itemProperties = itemType?.GetProperties().ToDictionary(p => p.Name, p => p);

        for (int i = options.DataStartRowIndex; i < table.Rows.Count; i++)
        {
            if (options.SkipBlankRows && IsRowBlank(table.Rows[i])) continue;

            var item = Activator.CreateInstance(itemType);
            for (int j = 0; j < table.Columns.Count; j++)
            {
                string header = table.Rows[options.DataStartRowIndex - 1][j].ToString();
                if (itemProperties != null && itemProperties.TryGetValue(header, out var property))
                {
                    var cellValue = table.Rows[i][j];
                    if (cellValue == DBNull.Value || (string.IsNullOrEmpty(cellValue.ToString()) && options.TreatEmptyAsNull))
                    {
                        property.SetValue(item, null);
                    }
                    else
                    {
                        SetPropertyValue(item, property, cellValue.ToString(), options);
                    }
                }
            }
            collection.Add(item);
        }

        return collection;
    }

    private T DeserializeSingleObject<T>(DataTable table, ExcelSerializerOptions options)
    {
        var obj = Activator.CreateInstance<T>();
        var properties = typeof(T).GetProperties().ToDictionary(p => p.Name, p => p);

        if (table.Rows.Count < options.DataStartRowIndex + 1) throw new Exception("Insufficient data for a single object.");

        for (int j = 0; j < table.Columns.Count; j++)
        {
            string header = table.Columns[j].ColumnName;
            if (properties.TryGetValue(header, out var property))
            {
                var cellValue = table.Rows[options.DataStartRowIndex][j];
                if (cellValue == DBNull.Value || (string.IsNullOrEmpty(cellValue.ToString()) && options.TreatEmptyAsNull))
                {
                    property.SetValue(obj, null);
                }
                else
                {
                    SetPropertyValue(obj, property, cellValue.ToString(), options);
                }
            }
        }

        return obj;
    }

    private void SetPropertyValue(object target, PropertyInfo property, string cellValue, ExcelSerializerOptions options)
    {
        if (property.PropertyType.IsArray || typeof(IEnumerable).IsAssignableFrom(property.PropertyType))
        {
            var elementType = property.PropertyType.IsArray
                ? property.PropertyType.GetElementType()
                : property.PropertyType.GetGenericArguments().FirstOrDefault();

            if (elementType != null)
            {
                var items = cellValue.Split(new[] { options.ArrayItemSeparator }, StringSplitOptions.None)
                                     .Select(item => Convert.ChangeType(item.Trim(), elementType))
                                     .ToArray();

                var typedArray = Array.CreateInstance(elementType, items.Length);
                items.CopyTo(typedArray, 0);
                property.SetValue(target, typedArray);
            }
            else
                property.SetValue(target, Convert.ChangeType(cellValue, property.PropertyType));
        }
        else
        {
            property.SetValue(target, Convert.ChangeType(cellValue, property.PropertyType));
        }
    }

    private bool IsRowBlank(DataRow row)
    {
        return row.ItemArray.All(field => field == DBNull.Value || string.IsNullOrWhiteSpace(field?.ToString()));
    }
}