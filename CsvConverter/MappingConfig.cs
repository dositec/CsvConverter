namespace CsvConverter
{
    public class DepersonalizationConfig
    {
        public string? PlaceholderFile { get; set; }
        public string? OutputSuffix { get; set; } = "_depersonalized";
        public List<DepersonalizationEntry>? Replacements { get; set; }
    }

    public class DepersonalizationEntry
    {
        public string Find { get; set; } = string.Empty;
        public string Replace { get; set; } = string.Empty;
        public bool WholeWordMatch { get; set; } = false;
    }

    public class DepersonalizationFile
    {
        public List<DepersonalizationEntry> Replacements { get; set; } = new List<DepersonalizationEntry>();
    }

    public class ReplacementConfig
    {
        public string? PlaceholderFile { get; set; }
        public string? OutputSuffix { get; set; } = "_replaced";
        public List<ReplacementEntry>? Replacements { get; set; }
    }

    public class ReplacementEntry
    {
        public string Find { get; set; } = string.Empty;
        public string Replace { get; set; } = string.Empty;
        public bool WholeWordMatch { get; set; } = false;
    }

    public class ReplacementFile
    {
        public List<ReplacementEntry> Replacements { get; set; } = new List<ReplacementEntry>();
    }

    public class PersonalizationConfig
    {
        public string? OutputSuffix { get; set; } = "_personalized";
    }

    public class MappingConfig
    {
        public string? Caption { get; set; }
        public string? CaptionBackgroundColor { get; set; }
        public string? HeaderBackgroundColor { get; set; }
        public List<string>? ZebraColors { get; set; }
        public string? CsvDelimiter { get; set; }
        public DepersonalizationConfig? Depersonalization { get; set; }
        public ReplacementConfig? Replacement { get; set; }
        public PersonalizationConfig? Personalization { get; set; }
        public List<ColumnMapping> Columns { get; set; } = new List<ColumnMapping>();
    }
}
