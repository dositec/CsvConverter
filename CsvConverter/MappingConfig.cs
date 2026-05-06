namespace CsvConverter
{
    public class MappingConfig
    {
        public string? Caption { get; set; }
        public string? CaptionBackgroundColor { get; set; }
        public string? HeaderBackgroundColor { get; set; }
        public List<string>? ZebraColors { get; set; }
        public List<ColumnMapping> Columns { get; set; } = new List<ColumnMapping>();
    }
}
