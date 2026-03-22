namespace WordTableProcessor.Models;

public enum DocumentCategory
{
    FER2020,
    TER2014,
    TER2010
}

public enum DocumentSubType
{
    Base,
    SSSC,
    SSEM
}

public class SmetaDocumentType
{
    public DocumentCategory Category { get; set; }
    public DocumentSubType SubType { get; set; }
    public string DisplayName { get; set; } = string.Empty;
    public string CodePattern { get; set; } = string.Empty;
    public int TableCount { get; set; } = 4;
    public string CsvSeparator { get; set; } = ";";
    public string[]? CsvHeaders { get; set; }

    public static List<SmetaDocumentType> GetAllTypes()
    {
        return new List<SmetaDocumentType>
        {
            new() { Category = DocumentCategory.FER2020, SubType = DocumentSubType.Base, DisplayName = "ФЕР-2020", CodePattern = @"^\d{2}-\d{2}-\d{3}-\d{2}.*$", TableCount = 4, CsvSeparator = ";", CsvHeaders = new[] { "TABLE_ID", "CODE", "PRICE1", "PRICE2" } },
            new() { Category = DocumentCategory.FER2020, SubType = DocumentSubType.SSSC, DisplayName = "ФССЦ-2020", CodePattern = @"^\d{3}-\d{4}.*$", TableCount = 10, CsvSeparator = ";", CsvHeaders = new[] { "TABLE_ID", "CODE", "NAME", "UNIT", "PRICE_OPT", "PRICE", "INDX" } },
            new() { Category = DocumentCategory.FER2020, SubType = DocumentSubType.SSEM, DisplayName = "ФСЭМ-2020", CodePattern = @"^\d{6}.*$", TableCount = 10, CsvSeparator = ";", CsvHeaders = new[] { "TABLE_ID", "CODE", "NAME", "UNIT", "PRICE", "OTM", "INDX" } },
            
            new() { Category = DocumentCategory.TER2014, SubType = DocumentSubType.Base, DisplayName = "ТЕР-2014", CodePattern = @"^\d{2}-\d{2}-\d{3}-\d{2}.*$", TableCount = 4, CsvSeparator = ";", CsvHeaders = new[] { "TABLE_ID", "CODE", "PRICE1", "PRICE2" } },
            new() { Category = DocumentCategory.TER2014, SubType = DocumentSubType.SSSC, DisplayName = "ТССЦ-2014", CodePattern = @"^\d{3}-\d{4}.*$", TableCount = 10, CsvSeparator = ";", CsvHeaders = new[] { "TABLE_ID", "CODE", "NAME", "UNIT", "PRICE_OPT", "PRICE", "INDX" } },
            new() { Category = DocumentCategory.TER2014, SubType = DocumentSubType.SSEM, DisplayName = "ТСЭМ-2014", CodePattern = @"^\d{6}.*$", TableCount = 10, CsvSeparator = ";", CsvHeaders = new[] { "TABLE_ID", "CODE", "NAME", "UNIT", "PRICE", "OTM", "INDX" } },
            
            new() { Category = DocumentCategory.TER2010, SubType = DocumentSubType.Base, DisplayName = "ТЕР-2010", CodePattern = @"^\d{2}-\d{2}-\d{3}-\d{2}.*$", TableCount = 4, CsvSeparator = ";", CsvHeaders = new[] { "TABLE_ID", "CODE", "PRICE1", "PRICE2" } },
            new() { Category = DocumentCategory.TER2010, SubType = DocumentSubType.SSSC, DisplayName = "ТССЦ-2010", CodePattern = @"^\d{3}-\d{4}.*$", TableCount = 10, CsvSeparator = ";", CsvHeaders = new[] { "TABLE_ID", "CODE", "NAME", "UNIT", "PRICE_OPT", "PRICE", "INDX" } },
            new() { Category = DocumentCategory.TER2010, SubType = DocumentSubType.SSEM, DisplayName = "ТСЭМ-2010", CodePattern = @"^\d{6}.*$", TableCount = 10, CsvSeparator = ";", CsvHeaders = new[] { "TABLE_ID", "CODE", "NAME", "UNIT", "PRICE", "OTM", "INDX" } }
        };
    }

    public static List<SmetaDocumentType> GetTypesByCategory(DocumentCategory category)
    {
        return GetAllTypes().Where(t => t.Category == category).ToList();
    }

    public static string GetCategoryDisplayName(DocumentCategory category)
    {
        return category switch
        {
            DocumentCategory.FER2020 => "ФЕР-2020 (ФССЦ, ФСЭМ)",
            DocumentCategory.TER2014 => "ТЕР-2014 (ТССЦ, ТСЭМ)",
            DocumentCategory.TER2010 => "ТЕР-2010 (ТССЦ, ТСЭМ)",
            _ => category.ToString()
        };
    }
}
