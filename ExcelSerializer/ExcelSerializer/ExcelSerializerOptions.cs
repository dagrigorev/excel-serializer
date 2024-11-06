namespace ExcelSerializer;

// Options Class
public class ExcelSerializerOptions
{
    public static ExcelSerializerOptions Default { get; } = new ExcelSerializerOptions();

    public int MaxFileSize { get; set; } = 10 * 1024 * 1024; // Default: 10 MB
    public int DataStartRowIndex { get; set; } = 1; // Start reading from the second row (0-indexed)
    public char ArrayItemSeparator { get; set; } = ','; // Default separator for array items
    public bool TreatEmptyAsNull { get; set; } = true; // Default: treat empty cells as null
    public bool SkipBlankRows { get; set; } = true; // Default: skip blank rows
}