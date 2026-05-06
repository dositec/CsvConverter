using ClosedXML.Excel;
using Microsoft.Extensions.Logging;
using YamlDotNet.Serialization.NamingConventions;
using YamlDotNet.Serialization;
using System.Text.Json.Serialization;
using System.Drawing;

namespace CsvConverter
{
    public class CsvColumnMapper
    {
        // Holds the column mapping loaded from YAML
        private List<ColumnMapping>? columnMappings;
        private readonly ILogger<CsvColumnMapper> logger;
        private readonly string? groupColumnInput;
        private readonly string? defaultValuesColumnInput;
        private readonly string? caption;
        private readonly string? captionBackgroundColor;
        private readonly string? headerBackgroundColor;
        private readonly List<string> zebraColors;
        private ColumnMapping? defaultValuesColumn;
        private int? groupColumnOutputIndex;
        
        public CsvColumnMapper(ILogger<CsvColumnMapper> logger, string configFile)
        {
            this.logger = logger;
            var config = LoadMappingConfig(configFile);
            columnMappings = config?.Columns;
            caption = config?.Caption;
            captionBackgroundColor = config?.CaptionBackgroundColor ?? "#4472C4";
            headerBackgroundColor = config?.HeaderBackgroundColor ?? "#D9E1F2";
            zebraColors = config?.ZebraColors ?? new List<string> { "#F2F2F2", "#FFFFFF" };
            groupColumnInput = DetermineGroupColumn();
            defaultValuesColumn = DetermineDefaultValuesColumn();
            defaultValuesColumnInput = defaultValuesColumn?.Input;
            if (!string.IsNullOrEmpty(groupColumnInput))
            {
                ReorderColumnsForGrouping();
            }
        }
        
        private string? DetermineGroupColumn()
        {
            if (columnMappings == null) return null;
            
            var groupColumns = columnMappings.Where(c => !string.IsNullOrEmpty(c.OutputGroup)).ToList();
            
            if (groupColumns.Count > 1)
            {
                logger.LogError("Only one column can have OutputGroup defined");
                return null;
            }
            
            return groupColumns.Count == 1 ? groupColumns[0].Input : null;
        }

        private ColumnMapping? DetermineDefaultValuesColumn()
        {
            if (columnMappings == null) return null;

            var defaultColumns = columnMappings.Where(c => c.DefaultValues != null && c.DefaultValues.Count > 0).ToList();
            if (defaultColumns.Count > 1)
            {
                logger.LogError("Only one column can have DefaultValues defined");
                return null;
            }

            return defaultColumns.Count == 1 ? defaultColumns[0] : null;
        }

        private void ReorderColumnsForGrouping()
        {
            if (columnMappings == null || string.IsNullOrEmpty(groupColumnInput)) return;

            // Find the group column and its original index
            var groupColumn = columnMappings.FirstOrDefault(c => c.Input == groupColumnInput);
            if (groupColumn == null) return;

            groupColumnOutputIndex = 1;

            // Move group column to position 1 and shift others
            columnMappings.Remove(groupColumn);
            groupColumn.OutputIndex = 1;
            columnMappings.Insert(0, groupColumn);

            // Renumber all other columns starting from 2
            for (int i = 1; i < columnMappings.Count; i++)
            {
                columnMappings[i].OutputIndex = i + 1;
            }

            logger.LogInformation($"Reordered columns: GroupColumn '{groupColumnInput}' moved to position 1");
        }
        
        public async Task<bool> ConvertAsync(string inputFile, string outputFile, CancellationToken token)
        {
            if (columnMappings == null)
            {
                logger.LogError("Column mapping is missing.");
                return false;
            }

            try
            {
                logger.LogInformation($"Starting conversion for file: {inputFile}");

                // Step 1: Read CSV file
                var csvLines = await ReadCsvFile(inputFile, token);
                if (csvLines == null) return false;

                // Step 2: Parse header row
                var csvColumnIndex = ParseCsvHeaders(csvLines[0]);
                if (csvColumnIndex == null) return false;

                // Step 3: Validate column mappings
                if (!ValidateColumnMappings(csvColumnIndex)) return false;

                // Step 4: Convert CSV to XLSX
                WriteToExcel(csvLines, csvColumnIndex, outputFile, token);

                logger.LogInformation($"CSV file successfully converted to XLSX at: {outputFile}");
                return true;
            }
            catch (OperationCanceledException)
            {
                logger.LogWarning($"Conversion canceled for file: {inputFile}");
                return false;
            }
            catch (Exception ex)
            {
                logger.LogError($"Error processing file {inputFile}: {ex.Message}");
                return false;
            }
        }

        private async Task<string[]?> ReadCsvFile(string inputFile, CancellationToken token)
        {
            try
            {
                var csvLines = await File.ReadAllLinesAsync(inputFile, token);
                logger.LogInformation($"Read {csvLines.Length} lines from {inputFile}");
                return csvLines;
            }
            catch (Exception ex)
            {
                logger.LogError($"Failed to read CSV file {inputFile}: {ex.Message}");
                return null;
            }
        }

        private Dictionary<string, int>? ParseCsvHeaders(string headerRow)
        {
            var csvHeaders = headerRow.Split(';');
            logger.LogInformation($"Parsed {csvHeaders.Length} columns in the header row");

            var csvColumnIndex = new Dictionary<string, int>();
            for (int i = 0; i < csvHeaders.Length; i++)
            {
                csvColumnIndex[csvHeaders[i]] = i;
            }

            return csvColumnIndex;
        }

        private bool ValidateColumnMappings(Dictionary<string, int> csvColumnIndex)
        {
            if (columnMappings == null)
            {
                logger.LogError("Column mapping is missing.");
                return false;
            }

            var missingColumns = columnMappings
                .Where(c => c != defaultValuesColumn && !csvColumnIndex.ContainsKey(c.Input))
                .Select(c => c.Input)
                .ToList();

            if (missingColumns.Count != 0)
            {
                logger.LogWarning($"Missing required columns in the CSV file: {string.Join(", ", missingColumns)}");
                return true;
            }

            if (defaultValuesColumn != null && !string.IsNullOrEmpty(defaultValuesColumn.DefaultValue))
            {
                logger.LogError("Column with DefaultValues cannot also specify DefaultValue.");
                return false;
            }

            if (defaultValuesColumn != null && !string.IsNullOrEmpty(groupColumnInput) && defaultValuesColumnInput == groupColumnInput)
            {
                logger.LogError("Column with DefaultValues cannot also be the OutputGroup column.");
                return false;
            }

            logger.LogInformation("All required columns are present in the CSV file.");
            return true;
        }

        private void WriteToExcel(string[] csvLines, Dictionary<string, int> csvColumnIndex, string outputFile, CancellationToken token)
        {
            if (columnMappings == null)
            {
                logger.LogError("Column mapping is missing.");
                return;
            }

            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Sheet1");

            // Write headers
            foreach (var mapping in columnMappings)
            {
                var headerName = mapping.Output ?? mapping.OutputGroup ?? "Unknown";
                worksheet.Cell(1, mapping.OutputIndex).Value = headerName;
            }
            logger.LogInformation("Headers written to the Excel file.");

            var headerRowIndex = string.IsNullOrEmpty(caption) ? 1 : 2;
            var maxOutputIndex = columnMappings.Max(c => c.OutputIndex);
            if (!string.IsNullOrEmpty(caption))
            {
                var captionRange = worksheet.Range(1, 1, 1, maxOutputIndex);
                captionRange.Merge();
                captionRange.Value = caption;
                captionRange.Style.Font.Bold = true;
                captionRange.Style.Font.FontSize = 14;
                captionRange.Style.Font.FontColor = XLColor.White;
                captionRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                captionRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                captionRange.Style.Fill.BackgroundColor = ParseZebraColor(captionBackgroundColor);
            }

            foreach (var mapping in columnMappings)
            {
                var headerName = mapping.Output ?? mapping.OutputGroup ?? "Unknown";
                worksheet.Cell(headerRowIndex, mapping.OutputIndex).Value = headerName;
            }

            int lastRowIndex;
            if (string.IsNullOrEmpty(groupColumnInput) && defaultValuesColumn == null)
            {
                // No grouping or row expansion - write data as is
                lastRowIndex = WriteDataWithoutGrouping(csvLines, csvColumnIndex, worksheet, headerRowIndex, token);
            }
            else
            {
                // Group data and/or expand rows based on DefaultValues
                lastRowIndex = WriteDataWithGroupingAndExpansion(csvLines, csvColumnIndex, worksheet, headerRowIndex, token);
            }

            logger.LogInformation("Data rows written to the Excel file.");

            var tableRange = worksheet.Range(1, 1, lastRowIndex, maxOutputIndex);
            tableRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            tableRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            tableRange.Style.Alignment.WrapText = true;
            tableRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            tableRange.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

            worksheet.Row(headerRowIndex).Style.Font.Bold = true;
            var headerRange = worksheet.Range(headerRowIndex, 1, headerRowIndex, maxOutputIndex);
            headerRange.Style.Fill.BackgroundColor = ParseZebraColor(headerBackgroundColor);
            headerRange.Style.Font.FontColor = XLColor.Black;

            for (int col = 1; col <= maxOutputIndex; col++)
            {
                double maxLength = 0;
                for (int row = 1; row <= lastRowIndex; row++)
                {
                    var cellValue = worksheet.Cell(row, col).GetString();
                    if (string.IsNullOrEmpty(cellValue))
                    {
                        continue;
                    }

                    var length = cellValue.Length;
                    if (length > maxLength)
                    {
                        maxLength = length;
                    }
                }

                if (maxLength > 0)
                {
                    var width = Math.Min(Math.Max(maxLength + 2, 10), 60);
                    worksheet.Column(col).Width = width;
                }
            }

            worksheet.Rows(1, lastRowIndex).AdjustToContents();

            var previousGroup = string.Empty;
            var fillOn = false;
            for (int row = headerRowIndex + 1; row <= lastRowIndex; row++)
            {
                var currentGroup = worksheet.Cell(row, 1).GetString();
                if (string.IsNullOrEmpty(currentGroup))
                {
                    currentGroup = previousGroup;
                }

                if (!string.Equals(currentGroup, previousGroup, StringComparison.Ordinal))
                {
                    fillOn = !fillOn;
                    previousGroup = currentGroup;
                }

                if (fillOn)
                {
                    worksheet.Range(row, 1, row, maxOutputIndex)
                        .Style.Fill.BackgroundColor = ParseZebraColor(zebraColors[0]);
                }
                else
                {
                    worksheet.Range(row, 1, row, maxOutputIndex)
                        .Style.Fill.BackgroundColor = ParseZebraColor(zebraColors.Count > 1 ? zebraColors[1] : "#FFFFFF");
                }
            }

            workbook.SaveAs(outputFile);
        }

        private int WriteDataWithoutGrouping(string[] csvLines, Dictionary<string, int> csvColumnIndex, IXLWorksheet worksheet, int headerRowIndex, CancellationToken token)
        {
            if (columnMappings == null) return headerRowIndex;

            int lastRowIndex = headerRowIndex;
            for (int rowIndex = 1; rowIndex < csvLines.Length; rowIndex++)
            {
                token.ThrowIfCancellationRequested();
                var csvRow = csvLines[rowIndex].Split(';');
                lastRowIndex = rowIndex + headerRowIndex;

                foreach (var mapping in columnMappings)
                {
                    if (csvColumnIndex.TryGetValue(mapping.Input, out int csvIndex))
                    {
                        var value = csvRow[csvIndex];
                        var processedValue = GetProcessedValue(value, mapping);
                        worksheet.Cell(lastRowIndex, mapping.OutputIndex).Value = processedValue;
                    }
                }
            }

            return lastRowIndex;
        }

        private string GetProcessedValue(string value, ColumnMapping mapping)
        {
            if (string.IsNullOrEmpty(value) && !string.IsNullOrEmpty(mapping.DefaultValue))
            {
                return mapping.DefaultValue;
            }

            if (!string.IsNullOrEmpty(mapping.Math))
            {
                value = ApplyMathOperation(value, mapping.Math);
            }

            return value;
        }

        private string ApplyMathOperation(string value, string mathOperation)
        {
            if (string.IsNullOrWhiteSpace(value) || !double.TryParse(value, out double numValue))
            {
                return value;
            }

            var parts = mathOperation.Trim().Split(new[] { ' ' }, System.StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length != 2)
            {
                return value;
            }

            var op = parts[0];
            if (!double.TryParse(parts[1], out double operand))
            {
                return value;
            }

            try
            {
                double result = op switch
                {
                    "+" => numValue + operand,
                    "-" => numValue - operand,
                    "*" => numValue * operand,
                    "/" => operand != 0 ? numValue / operand : numValue,
                    _ => numValue
                };

                return result.ToString("G15");
            }
            catch
            {
                return value;
            }
        }

        private int WriteDataWithGroupingAndExpansion(string[] csvLines, Dictionary<string, int> csvColumnIndex, IXLWorksheet worksheet, int headerRowIndex, CancellationToken token)
        {
            if (columnMappings == null) return headerRowIndex;

            int? groupColumnCsvIndex = null;
            if (!string.IsNullOrEmpty(groupColumnInput))
            {
                if (!csvColumnIndex.TryGetValue(groupColumnInput, out int index))
                {
                    logger.LogError($"Group column '{groupColumnInput}' not found in CSV");
                    return headerRowIndex;
                }
                groupColumnCsvIndex = index;
            }

            // Parse all data rows into groups (if grouping) or flat order
            var groupedRows = new Dictionary<string, List<(string[] CsvRow, int ExpandCount)>>();
            var groupOrder = new List<string>();

            for (int rowIndex = 1; rowIndex < csvLines.Length; rowIndex++)
            {
                token.ThrowIfCancellationRequested();
                var csvRow = csvLines[rowIndex].Split(';');
                var groupKey = groupColumnCsvIndex.HasValue ? csvRow[groupColumnCsvIndex.Value] : string.Empty;

                if (!groupedRows.ContainsKey(groupKey))
                {
                    groupedRows[groupKey] = new List<(string[] CsvRow, int ExpandCount)>();
                    groupOrder.Add(groupKey);
                }

                var expandCount = defaultValuesColumn != null ? defaultValuesColumn.DefaultValues!.Count : 1;
                groupedRows[groupKey].Add((csvRow, expandCount));
            }

            int outputRowIndex = headerRowIndex + 1;
            var groupMergeRanges = new List<(int startRow, int endRow)>();
            var rowMergeRanges = new List<(int startRow, int endRow, int outputIndex)>();

            foreach (var groupKey in groupOrder)
            {
                token.ThrowIfCancellationRequested();
                var rows = groupedRows[groupKey];
                int groupStartRow = outputRowIndex;

                foreach (var (csvRow, expandCount) in rows)
                {
                    int rowStartRow = outputRowIndex;

                    for (int partIndex = 0; partIndex < expandCount; partIndex++)
                    {
                        foreach (var mapping in columnMappings)
                        {
                            if (!string.IsNullOrEmpty(mapping.OutputGroup))
                            {
                                if (outputRowIndex == groupStartRow)
                                {
                                    worksheet.Cell(outputRowIndex, mapping.OutputIndex).Value = groupKey;
                                }
                            }
                            else if (defaultValuesColumn != null && mapping.Input == defaultValuesColumn.Input)
                            {
                                worksheet.Cell(outputRowIndex, mapping.OutputIndex).Value = defaultValuesColumn.DefaultValues![partIndex];
                            }
                            else if (!csvColumnIndex.TryGetValue(mapping.Input, out int csvIndex))
                            {
                                continue;
                            }
                            else
                            {
                                var value = csvRow[csvIndex];
                                var processedValue = GetProcessedValue(value, mapping);
                                worksheet.Cell(outputRowIndex, mapping.OutputIndex).Value = processedValue;
                            }
                        }

                        outputRowIndex++;
                    }

                    if (outputRowIndex - rowStartRow > 1)
                    {
                        foreach (var mapping in columnMappings)
                        {
                            if (defaultValuesColumn != null && mapping.Input == defaultValuesColumn.Input)
                            {
                                continue;
                            }

                            if (!string.IsNullOrEmpty(mapping.OutputGroup))
                            {
                                continue;
                            }

                            rowMergeRanges.Add((rowStartRow, outputRowIndex - 1, mapping.OutputIndex));
                        }
                    }
                }

                if (outputRowIndex - groupStartRow > 1 && groupColumnOutputIndex.HasValue)
                {
                    groupMergeRanges.Add((groupStartRow, outputRowIndex - 1));
                }
            }

            foreach (var (startRow, endRow, outputIndex) in rowMergeRanges)
            {
                token.ThrowIfCancellationRequested();
                var rangeToMerge = worksheet.Range(startRow, outputIndex, endRow, outputIndex);
                rangeToMerge.Merge();
            }

            if (groupColumnOutputIndex.HasValue)
            {
                foreach (var (startRow, endRow) in groupMergeRanges)
                {
                    token.ThrowIfCancellationRequested();
                    var rangeToMerge = worksheet.Range(startRow, groupColumnOutputIndex.Value, endRow, groupColumnOutputIndex.Value);
                    rangeToMerge.Merge();
                    rangeToMerge.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    rangeToMerge.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
                }

                logger.LogInformation($"Merged {groupMergeRanges.Count} cell ranges for group column");
            }

            return Math.Max(headerRowIndex, outputRowIndex - 1);
        }

        private MappingConfig? LoadMappingConfig(string yamlFilePath)
        {
            try
            {
                var deserializer = new DeserializerBuilder()
                    .WithNamingConvention(NullNamingConvention.Instance)
                    .Build();
 
                logger.LogInformation($"Loading YAML configuration from {yamlFilePath}");
                using var reader = new StreamReader(yamlFilePath);
                var config = deserializer.Deserialize<MappingConfig>(reader);
                if (config == null)
                {
                    logger.LogError("YAML configuration deserialized to null.");
                    return null;
                }

                return config;
            }
            catch (Exception ex)
            {
                logger.LogError($"Error loading YAML file: {ex.Message}");
                return null;
            }
        }

        private XLColor ParseZebraColor(string color)
        {
            if (string.IsNullOrWhiteSpace(color))
            {
                return XLColor.NoColor;
            }

            color = color.Trim();
            if (color.StartsWith("#"))
            {
                try
                {
                    var r = Convert.ToInt32(color.Substring(1, 2), 16);
                    var g = Convert.ToInt32(color.Substring(3, 2), 16);
                    var b = Convert.ToInt32(color.Substring(5, 2), 16);
                    return XLColor.FromArgb(r, g, b);
                }
                catch
                {
                    return XLColor.NoColor;
                }
            }

            return XLColor.FromName(color);
        }
    }
}
