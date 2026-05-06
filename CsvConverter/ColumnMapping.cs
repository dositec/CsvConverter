namespace CsvConverter
{
    public class ColumnMapping
    {
        public string Input { get; set; } = string.Empty;
        public int OutputIndex { get; set; } = 0;
        public string? Output { get; set; }
        public string? OutputGroup { get; set; }
        public string? DefaultValue { get; set; }
        public List<string>? DefaultValues { get; set; }
        public string? Math { get; set; }
    }
}
