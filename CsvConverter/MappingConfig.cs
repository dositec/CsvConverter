namespace CsvConverter
{
    public class MappingConfig
    {
        public string? Caption { get; set; }
        public List<ColumnMapping> Columns { get; set; } = new List<ColumnMapping>();
    }
}
