using ClosedXML.Excel;
using Microsoft.Extensions.Logging;
using YamlDotNet.Serialization.NamingConventions;
using YamlDotNet.Serialization;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using System.Drawing;
using System.IO;

namespace CsvConverter
{
    public class CsvColumnMapper
    {
        // Holds the column mapping loaded from YAML
        private List<ColumnMapping>? columnMappings;
        private readonly ILogger<CsvColumnMapper> logger;
        private readonly string? groupColumnInput;
        private readonly string? defaultValuesColumnInput;
        private readonly string? aiValuesColumnInput;
        private readonly string csvDelimiter;
        private readonly string? caption;
        private readonly string? captionBackgroundColor;
        private readonly string? headerBackgroundColor;
        private readonly List<string> zebraColors;
        private readonly List<DepersonalizationEntry> depersonalizationEntries;
        private readonly DepersonalizationConfig? depersonalizationConfig;
        private readonly List<ReplacementEntry> replacementEntries;
        private readonly ReplacementConfig? replacementConfig;
        private readonly PersonalizationConfig? personalizationConfig;
        private IAIProvider? aiProvider;
        private readonly string configDirectory;
        private readonly string configFilePath;
        private readonly AIConfig? aiConfig;
        private Task? aiProviderInitializationTask;
        private ColumnMapping? defaultValuesColumn;
        private ColumnMapping? aiValuesColumn;
        private int? groupColumnOutputIndex;
        
        public CsvColumnMapper(ILogger<CsvColumnMapper> logger, string configFile, string? depersonalizationFilePath = null)
        {
            this.logger = logger;
            configFilePath = configFile;
            var config = LoadMappingConfig(configFile);
            logger.LogError($"Config loaded: {(config != null ? "not null" : "null")}, Columns count: {config?.Columns?.Count ?? 0}, AI count: {config?.AI?.Count ?? 0}");
            columnMappings = config?.Columns;
            caption = config?.Caption;
            captionBackgroundColor = config?.CaptionBackgroundColor ?? "#4472C4";
            headerBackgroundColor = config?.HeaderBackgroundColor ?? "#D9E1F2";
            zebraColors = config?.ZebraColors ?? new List<string> { "#F2F2F2", "#FFFFFF" };
            csvDelimiter = DetermineCsvDelimiter(config?.CsvDelimiter);
            depersonalizationConfig = config?.Depersonalization;
            replacementConfig = config?.Replacement;
            personalizationConfig = config?.Personalization;
            aiConfig = config?.AI?.FirstOrDefault();
            if (aiConfig != null)
            {
                logger.LogError($"AI config loaded in constructor: Name='{aiConfig.Name}', URL='{aiConfig.Url}', Model='{aiConfig.Model}'");
            }
            else
            {
                logger.LogError("No AI config found in configuration file in constructor.");
            }
            aiProvider = null; // Will be initialized asynchronously via InitializeAIProviderAsync
            configDirectory = Path.GetDirectoryName(configFile) ?? string.Empty;

            var rootDepersonalizationEntries = config?.Depersonalization?.Replacements ?? new List<DepersonalizationEntry>();
            var rootReplacementEntries = config?.Replacement?.Replacements ?? new List<ReplacementEntry>();

            depersonalizationEntries = LoadDepersonalizationEntries(depersonalizationFilePath)
                .Concat(rootDepersonalizationEntries)
                .ToList();
            replacementEntries = LoadReplacementEntries()
                .Concat(rootReplacementEntries)
                .ToList();
            groupColumnInput = DetermineGroupColumn();
            defaultValuesColumn = DetermineDefaultValuesColumn();
            defaultValuesColumnInput = defaultValuesColumn?.Input;
            aiValuesColumn = DetermineAIValuesColumn();
            aiValuesColumnInput = aiValuesColumn?.Input;
            if (!string.IsNullOrEmpty(groupColumnInput))
            {
                ReorderColumnsForGrouping();
            }
        }

        public void StartInitializeAIProviderAsync(CancellationToken token = default)
        {
            if (aiProviderInitializationTask != null)
            {
                logger.LogInformation("AI provider initialization already in progress.");
                return; // Already initialized or in progress
            }
                
            logger.LogInformation("Starting AI provider initialization...");
            aiProviderInitializationTask = InitializeAIProviderInternalAsync(token);
        }

        private async Task InitializeAIProviderInternalAsync(CancellationToken token)
        {
            if (aiConfig != null)
            {
                logger.LogInformation($"Initializing AI provider with config: Name='{aiConfig.Name}', URL='{aiConfig.Url}', Model='{aiConfig.Model}'");
                try
                {
                    aiProvider = await AIProviderFactory.CreateProviderAsync(aiConfig, token);
                    if (aiProvider == null)
                    {
                        logger.LogWarning($"AI configuration is present but AIProvider could not be created. Check if Ollama service is running and accessible at {aiConfig?.Url}. Verify the model '{aiConfig?.Model}' exists in Ollama.");
                    }
                    else
                    {
                        logger.LogError("AIProvider initialized successfully");
                    }
                }
                catch (Exception ex)
                {
                    logger.LogError($"Error initializing AIProvider: {ex.Message}");
                }
            }
            else
            {
                logger.LogWarning("AI config is null. AI provider will not be initialized.");
            }
        }

        private async Task EnsureAIProviderInitializedAsync(CancellationToken token)
        {
            if (aiProviderInitializationTask != null)
            {
                logger.LogInformation("Waiting for AI provider initialization to complete...");
                try
                {
                    await aiProviderInitializationTask;
                    logger.LogInformation("AI provider initialization task completed.");
                    logger.LogError($"AI provider state after initialization: {(aiProvider != null ? "Initialized" : "NULL")}");
                }
                catch (Exception ex)
                {
                    logger.LogError($"AI provider initialization task failed: {ex.Message}");
                    // Error already logged in InitializeAIProviderInternalAsync
                }
            }
            else
            {
                logger.LogWarning("AI provider initialization task is null. AI provider may not have been started.");
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

        private ColumnMapping? DetermineAIValuesColumn()
        {
            if (columnMappings == null) return null;

            var aiColumns = columnMappings.Where(c => c.AIValues != null && c.AIValues.Count > 0).ToList();
            if (aiColumns.Count > 1)
            {
                logger.LogError("Only one column can have AIValues defined");
                return null;
            }

            return aiColumns.Count == 1 ? aiColumns[0] : null;
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
            var mappingCount = columnMappings?.Count ?? 0;
            if (columnMappings == null || mappingCount == 0)
            {
                logger.LogError($"Column mapping is missing for config '{configFilePath}'. Loaded column count: {mappingCount}. Ensure the YAML file contains a valid 'Columns:' section.");
                return false;
            }

            // Ensure AI provider is initialized if needed
            if (aiValuesColumn != null)
            {
                await EnsureAIProviderInitializedAsync(token);
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
                await WriteToExcel(csvLines, csvColumnIndex, outputFile, token);

                if (replacementEntries.Any())
                {
                    var replacedOutputFile = GetReplacementOutputFilePath(outputFile);
                    await WriteToExcel(csvLines, csvColumnIndex, replacedOutputFile, token, false, true);
                    logger.LogInformation($"Replaced CSV file created at: {replacedOutputFile}");

                    if (depersonalizationEntries.Any())
                    {
                        var depersonalizedOutputFile = GetDepersonalizedOutputFilePath(replacedOutputFile);
                        await WriteToExcel(csvLines, csvColumnIndex, depersonalizedOutputFile, token, true, true);
                        logger.LogInformation($"Depersonalized CSV file created at: {depersonalizedOutputFile}");
                    }
                }
                else if (depersonalizationEntries.Any())
                {
                    var depersonalizedOutputFile = GetDepersonalizedOutputFilePath(outputFile);
                    await WriteToExcel(csvLines, csvColumnIndex, depersonalizedOutputFile, token, true);
                    logger.LogInformation($"Depersonalized CSV file created at: {depersonalizedOutputFile}");
                }

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

        public async Task<bool> PersonalizeAsync(string inputFile, string outputFile, CancellationToken token)
        {
            if (!depersonalizationEntries.Any())
            {
                logger.LogError("No depersonalization entries loaded for personalization.");
                return false;
            }

            try
            {
                logger.LogInformation($"Starting personalization for file: {inputFile}");

                using var workbook = new XLWorkbook(inputFile);
                var worksheet = workbook.Worksheets.First();

                // Apply reverse replacements
                foreach (var row in worksheet.RowsUsed())
                {
                    foreach (var cell in row.CellsUsed())
                    {
                        if (cell.DataType == XLDataType.Text)
                        {
                            var originalValue = cell.GetString();
                            var personalizedValue = ApplyPersonalization(originalValue);
                            cell.Value = personalizedValue;
                        }
                    }
                }

                workbook.SaveAs(outputFile);
                logger.LogInformation($"Personalized file saved to: {outputFile}");
                return true;
            }
            catch (Exception ex)
            {
                logger.LogError($"Error personalizing file {inputFile}: {ex.Message}");
                return false;
            }
        }

        private string ApplyPersonalization(string value)
        {
            if (depersonalizationEntries == null || !depersonalizationEntries.Any() || string.IsNullOrEmpty(value))
            {
                return value;
            }

            var result = value;

            foreach (var replacement in depersonalizationEntries)
            {
                if (string.IsNullOrEmpty(replacement.Replace))
                {
                    continue;
                }

                // Reverse: replace the placeholder with the original
                if (replacement.WholeWordMatch)
                {
                    var pattern = $"\\b{Regex.Escape(replacement.Replace)}\\b";
                    result = Regex.Replace(result, pattern, _ => replacement.Find);
                }
                else
                {
                    result = result.Replace(replacement.Replace, replacement.Find, StringComparison.Ordinal);
                }
            }

            return result;
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
            var csvHeaders = headerRow.Split(new[] { csvDelimiter }, StringSplitOptions.None);
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
                .Where(c => !string.IsNullOrEmpty(c.Input) && c != defaultValuesColumn && c != aiValuesColumn && !csvColumnIndex.ContainsKey(c.Input))
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

            if (aiValuesColumn != null && !string.IsNullOrEmpty(aiValuesColumn.DefaultValue))
            {
                logger.LogError("Column with AIValues cannot also specify DefaultValue.");
                return false;
            }

            if (defaultValuesColumn != null && aiValuesColumn != null)
            {
                logger.LogError("Cannot have both DefaultValues and AIValues columns.");
                return false;
            }

            if (defaultValuesColumn != null && !string.IsNullOrEmpty(groupColumnInput) && defaultValuesColumnInput == groupColumnInput)
            {
                logger.LogError("Column with DefaultValues cannot also be the OutputGroup column.");
                return false;
            }

            if (aiValuesColumn != null && !string.IsNullOrEmpty(groupColumnInput) && aiValuesColumnInput == groupColumnInput)
            {
                logger.LogError("Column with AIValues cannot also be the OutputGroup column.");
                return false;
            }

            if (aiValuesColumn != null && aiProvider == null)
            {
                logger.LogWarning("AIValues column is configured but AI provider is not available. Conversion will continue using prompt text instead of AI-generated content. Make sure Ollama is running and configured correctly in Config.yaml. Check that the model specified in Config.yaml exists in Ollama (available models: llama3.2-vision:11b, qwen3.6:35b-a3b-q4_K_M, jeffh/intfloat-multilingual-e5-large:q8_0, qwen3-coder:30b).");
                // Continue conversion with fallback (prompt text will be used instead of AI-generated content)
            }

            logger.LogInformation("All required columns are present in the CSV file.");
            return true;
        }

        private async Task WriteToExcel(string[] csvLines, Dictionary<string, int> csvColumnIndex, string outputFile, CancellationToken token, bool depersonalize = false, bool applyReplacement = false)
        {
            if (columnMappings == null)
            {
                logger.LogError("Column mapping is missing.");
                return;
            }

            // Set AI provider log directory to the same folder as the output file
            var outputDir = Path.GetDirectoryName(outputFile);
            if (string.IsNullOrEmpty(outputDir))
            {
                outputDir = Directory.GetCurrentDirectory();
                logger.LogInformation($"Output file has no directory, using current directory: {outputDir}");
            }
            
            // Ensure directory exists
            try
            {
                Directory.CreateDirectory(outputDir);
                AIProviderFactory.LogDirectory = outputDir;
                logger.LogInformation($"AI provider logs will be written to: {outputDir}");
            }
            catch (Exception ex)
            {
                logger.LogWarning($"Failed to set AI log directory to {outputDir}: {ex.Message}. Using default location.");
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
            if (string.IsNullOrEmpty(groupColumnInput) && defaultValuesColumn == null && aiValuesColumn == null)
            {
                // No grouping or row expansion - write data as is
                lastRowIndex = WriteDataWithoutGrouping(csvLines, csvColumnIndex, worksheet, headerRowIndex, token, depersonalize, applyReplacement);
            }
            else
            {
                // Group data and/or expand rows based on DefaultValues or AIValues
                lastRowIndex = await WriteDataWithGroupingAndExpansion(csvLines, csvColumnIndex, worksheet, headerRowIndex, token, depersonalize, applyReplacement);
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

        private int WriteDataWithoutGrouping(string[] csvLines, Dictionary<string, int> csvColumnIndex, IXLWorksheet worksheet, int headerRowIndex, CancellationToken token, bool depersonalize, bool applyReplacement)
        {
            if (columnMappings == null) return headerRowIndex;

            int lastRowIndex = headerRowIndex;
            for (int rowIndex = 1; rowIndex < csvLines.Length; rowIndex++)
            {
                token.ThrowIfCancellationRequested();
                var csvRow = csvLines[rowIndex].Split(new[] { csvDelimiter }, StringSplitOptions.None);
                lastRowIndex = rowIndex + headerRowIndex;

                foreach (var mapping in columnMappings)
                {
                    string cellValue;
                    if (string.IsNullOrEmpty(mapping.Input))
                    {
                        // New column with fixed values
                        cellValue = mapping.DefaultValue ?? mapping.DefaultValues?.FirstOrDefault() ?? "";
                    }
                    else if (csvColumnIndex.TryGetValue(mapping.Input, out int csvIndex))
                    {
                        var value = csvRow[csvIndex];
                        cellValue = GetProcessedValue(value, mapping);
                    }
                    else
                    {
                        continue;
                    }

                    if (applyReplacement)
                    {
                        cellValue = ApplyReplacement(cellValue);
                    }

                    if (depersonalize)
                    {
                        cellValue = ApplyDepersonalization(cellValue);
                    }

                    worksheet.Cell(lastRowIndex, mapping.OutputIndex).Value = cellValue;
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

            if (!string.IsNullOrEmpty(mapping.MathRound))
            {
                value = ApplyMathRoundOperation(value, mapping.MathRound);
            }
            else if (!string.IsNullOrEmpty(mapping.Math))
            {
                value = ApplyMathOperation(value, mapping.Math);
            }

            return value;
        }

        private string ApplyMathRoundOperation(string value, string mathOperation)
        {
            if (string.IsNullOrWhiteSpace(value) || !double.TryParse(value, out double numValue))
            {
                return value;
            }

            var resultText = ApplyMathOperation(value, mathOperation);
            if (string.IsNullOrWhiteSpace(resultText) || !double.TryParse(resultText, out double resultValue))
            {
                return value;
            }

            return Math.Ceiling(resultValue).ToString("G15");
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

        private async Task<int> WriteDataWithGroupingAndExpansion(string[] csvLines, Dictionary<string, int> csvColumnIndex, IXLWorksheet worksheet, int headerRowIndex, CancellationToken token, bool depersonalize, bool applyReplacement)
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

            // Helper function to build context from entire row
            string BuildRowContext(string[] row)
            {
                var parts = new List<string>();
                foreach (var kvp in csvColumnIndex)
                {
                    var columnName = kvp.Key;
                    var columnIndex = kvp.Value;
                    if (columnIndex < row.Length)
                    {
                        var value = row[columnIndex];
                        if (!string.IsNullOrWhiteSpace(value))
                        {
                            parts.Add($"{columnName}: {value}");
                        }
                    }
                }
                return string.Join("; ", parts);
            }

            // Parse all data rows into groups (if grouping) or flat order
            var groupedRows = new Dictionary<string, List<(string[] CsvRow, int ExpandCount)>>();
            var groupOrder = new List<string>();

            for (int rowIndex = 1; rowIndex < csvLines.Length; rowIndex++)
            {
                token.ThrowIfCancellationRequested();
                var csvRow = csvLines[rowIndex].Split(new[] { csvDelimiter }, StringSplitOptions.None);
                var groupKey = groupColumnCsvIndex.HasValue ? csvRow[groupColumnCsvIndex.Value] : string.Empty;

                if (!groupedRows.ContainsKey(groupKey))
                {
                    groupedRows[groupKey] = new List<(string[] CsvRow, int ExpandCount)>();
                    groupOrder.Add(groupKey);
                }

                var expandCount = aiValuesColumn != null ? aiValuesColumn.AIValues!.Count : defaultValuesColumn != null ? defaultValuesColumn.DefaultValues!.Count : 1;
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
                                    var value = groupKey;
                                    if (applyReplacement)
                                    {
                                        value = ApplyReplacement(value);
                                    }
                                    if (depersonalize)
                                    {
                                        value = ApplyDepersonalization(value);
                                    }
                                    var cell = worksheet.Cell(outputRowIndex, mapping.OutputIndex);
                                    cell.Value = value;
                                    cell.Style.Alignment.WrapText = true;
                                    cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Top;
                                }
                            }
                            else if (aiValuesColumn != null && mapping == aiValuesColumn)
                            {
                                var prompt = aiValuesColumn.AIValues![partIndex];
                                // Enhance prompt with context from the entire row
                                var rowContext = BuildRowContext(csvRow);
                                if (!string.IsNullOrWhiteSpace(rowContext))
                                {
                                    prompt = $"{prompt}. Контекст: {rowContext}";
                                    logger.LogInformation($"Enhanced AI prompt with full row context: '{prompt}'");
                                }
                                string value;
                                if (aiProvider != null)
                                {
                                    try
                                    {
                                        value = await aiProvider.GenerateAsync(prompt, token);
                                        logger.LogInformation($"AI generation succeeded for prompt '{prompt}' -> '{value.Substring(0, Math.Min(value.Length, 100))}...'");
                                    }
                                    catch (Exception ex)
                                    {
                                        logger.LogError($"AI generation failed for prompt '{prompt}': {ex.Message}. Using prompt text instead.");
                                        value = prompt;
                                    }
                                }
                                else
                                {
                                    logger.LogWarning($"AIProvider is null. Using prompt text instead of AI response: '{prompt}'");
                                    value = prompt;
                                }
                                
                                if (applyReplacement)
                                {
                                    value = ApplyReplacement(value);
                                }
                                if (depersonalize)
                                {
                                    value = ApplyDepersonalization(value);
                                }
                                worksheet.Cell(outputRowIndex, mapping.OutputIndex).Value = value;
                            }
                            else if (defaultValuesColumn != null && mapping == defaultValuesColumn)
                            {
                                var value = defaultValuesColumn.DefaultValues![partIndex];
                                if (applyReplacement)
                                {
                                    value = ApplyReplacement(value);
                                }
                                if (depersonalize)
                                {
                                    value = ApplyDepersonalization(value);
                                }
                                worksheet.Cell(outputRowIndex, mapping.OutputIndex).Value = value;
                            }
                            else if (string.IsNullOrEmpty(mapping.Input))
                            {
                                // New column with fixed values
                                var value = mapping.DefaultValues != null
                                    ? mapping.DefaultValues[partIndex % mapping.DefaultValues.Count]
                                    : mapping.DefaultValue ?? string.Empty;
                                if (applyReplacement)
                                {
                                    value = ApplyReplacement(value);
                                }
                                if (depersonalize)
                                {
                                    value = ApplyDepersonalization(value);
                                }
                                worksheet.Cell(outputRowIndex, mapping.OutputIndex).Value = value;
                            }
                            else if (!csvColumnIndex.TryGetValue(mapping.Input, out int csvIndex))
                            {
                                continue;
                            }
                            else
                            {
                                var value = csvRow[csvIndex];
                                var processedValue = GetProcessedValue(value, mapping);
                                if (applyReplacement)
                                {
                                    processedValue = ApplyReplacement(processedValue);
                                }
                                if (depersonalize)
                                {
                                    processedValue = ApplyDepersonalization(processedValue);
                                }
                                worksheet.Cell(outputRowIndex, mapping.OutputIndex).Value = processedValue;
                            }
                        }

                        outputRowIndex++;
                    }

                    if (outputRowIndex - rowStartRow > 1)
                    {
                        foreach (var mapping in columnMappings)
                        {
                            if (aiValuesColumn != null && mapping == aiValuesColumn)
                            {
                                continue;
                            }

                            if (defaultValuesColumn != null && mapping == defaultValuesColumn)
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

            // Auto-adjust row heights for all data rows to fit AI responses
            if (outputRowIndex - 1 > headerRowIndex)
            {
                worksheet.Rows(headerRowIndex + 1, outputRowIndex - 1).AdjustToContents();
            }

            return Math.Max(headerRowIndex, outputRowIndex - 1);
        }

        private MappingConfig? LoadMappingConfig(string yamlFilePath)
        {
            try
            {
                var deserializer = new DeserializerBuilder()
                    .WithNamingConvention(NullNamingConvention.Instance)
                    .IgnoreUnmatchedProperties()
                    .Build();

                logger.LogInformation($"Loading YAML configuration from {yamlFilePath}");
                using var reader = new StreamReader(yamlFilePath);
                var config = deserializer.Deserialize<MappingConfig>(reader);
                if (config == null)
                {
                    logger.LogError("YAML configuration deserialized to null.");
                    return null;
                }

                logger.LogInformation($"Loaded YAML configuration with {config.Columns?.Count ?? 0} columns from {yamlFilePath}");
                
                // Log AI configuration details
                if (config.AI != null)
                {
                    logger.LogError($"AI configuration count: {config.AI.Count}");
                    if (config.AI.Count > 0)
                    {
                        var aiConfig = config.AI.First();
                        logger.LogError($"AI configuration loaded: Name='{aiConfig.Name}', URL='{aiConfig.Url}', Model='{aiConfig.Model}', TimeoutSeconds={aiConfig.TimeoutSeconds}");
                    }
                    else
                    {
                        logger.LogError("AI configuration list is empty.");
                    }
                }
                else
                {
                    logger.LogError("AI configuration is null in YAML file.");
                }
                
                return config;
            }
            catch (Exception ex)
            {
                logger.LogError($"Error loading YAML file: {ex.Message}");
                return null;
            }
        }

        private List<DepersonalizationEntry> LoadDepersonalizationEntries(string? depersonalizationFilePath)
        {
            var placeholderFile = GetDepersonalizationPlaceholderFilePath(depersonalizationFilePath);
            if (string.IsNullOrEmpty(placeholderFile))
            {
                logger.LogInformation("No depersonalization placeholder file configured.");
                return new List<DepersonalizationEntry>();
            }

            if (!File.Exists(placeholderFile))
            {
                logger.LogInformation($"Depersonalization placeholder file not found: {placeholderFile}");
                return new List<DepersonalizationEntry>();
            }

            try
            {
                var deserializer = new DeserializerBuilder()
                    .WithNamingConvention(NullNamingConvention.Instance)
                    .Build();

                logger.LogInformation($"Loading depersonalization placeholders from {placeholderFile}");
                using var reader = new StreamReader(placeholderFile);
                var file = deserializer.Deserialize<DepersonalizationFile>(reader);
                return file?.Replacements ?? new List<DepersonalizationEntry>();
            }
            catch (Exception ex)
            {
                logger.LogError($"Failed to load depersonalization placeholders: {ex.Message}");
                return new List<DepersonalizationEntry>();
            }
        }

        private List<ReplacementEntry> LoadReplacementEntries()
        {
            var placeholderFile = GetReplacementPlaceholderFilePath();
            if (string.IsNullOrEmpty(placeholderFile))
            {
                logger.LogInformation("No replacement placeholder file configured.");
                return new List<ReplacementEntry>();
            }

            if (!File.Exists(placeholderFile))
            {
                logger.LogInformation($"Replacement placeholder file not found: {placeholderFile}");
                return new List<ReplacementEntry>();
            }

            try
            {
                var deserializer = new DeserializerBuilder()
                    .WithNamingConvention(NullNamingConvention.Instance)
                    .Build();

                logger.LogInformation($"Loading replacement placeholders from {placeholderFile}");
                using var reader = new StreamReader(placeholderFile);
                var file = deserializer.Deserialize<ReplacementFile>(reader);
                return file?.Replacements ?? new List<ReplacementEntry>();
            }
            catch (Exception ex)
            {
                logger.LogError($"Failed to load replacement placeholders: {ex.Message}");
                return new List<ReplacementEntry>();
            }
        }

        private string GetDepersonalizationPlaceholderFilePath(string? depersonalizationFilePath)
        {
            if (!string.IsNullOrEmpty(depersonalizationFilePath))
            {
                return Path.IsPathRooted(depersonalizationFilePath)
                    ? depersonalizationFilePath
                    : Path.Combine(configDirectory, depersonalizationFilePath);
            }

            var placeholderFile = depersonalizationConfig?.PlaceholderFile ?? "depersonalization.yaml";
            return Path.Combine(configDirectory, placeholderFile);
        }

        private string GetReplacementPlaceholderFilePath()
        {
            var placeholderFile = replacementConfig?.PlaceholderFile ?? "replacements.yaml";
            return Path.Combine(configDirectory, placeholderFile);
        }

        private string GetDepersonalizedOutputFilePath(string originalOutputFile)
        {
            var suffix = depersonalizationConfig?.OutputSuffix ?? "_depersonalized";
            var directory = Path.GetDirectoryName(originalOutputFile) ?? string.Empty;
            var fileName = Path.GetFileNameWithoutExtension(originalOutputFile);
            var extension = Path.GetExtension(originalOutputFile);
            return Path.Combine(directory, fileName + suffix + extension);
        }

        private string GetReplacementOutputFilePath(string originalOutputFile)
        {
            var suffix = replacementConfig?.OutputSuffix ?? "_replaced";
            var directory = Path.GetDirectoryName(originalOutputFile) ?? string.Empty;
            var fileName = Path.GetFileNameWithoutExtension(originalOutputFile);
            var extension = Path.GetExtension(originalOutputFile);
            return Path.Combine(directory, fileName + suffix + extension);
        }

        public string GetPersonalizedOutputFilePath(string originalOutputFile)
        {
            var suffix = personalizationConfig?.OutputSuffix ?? "_personalized";
            var directory = Path.GetDirectoryName(originalOutputFile) ?? string.Empty;
            var fileName = Path.GetFileNameWithoutExtension(originalOutputFile);
            var extension = Path.GetExtension(originalOutputFile);
            return Path.Combine(directory, fileName + suffix + extension);
        }

        private string ApplyDepersonalization(string value)
        {
            if (depersonalizationEntries == null || !depersonalizationEntries.Any() || string.IsNullOrEmpty(value))
            {
                return value;
            }

            var result = value;

            foreach (var replacement in depersonalizationEntries)
            {
                if (string.IsNullOrEmpty(replacement.Find))
                {
                    continue;
                }

                if (replacement.WholeWordMatch)
                {
                    var pattern = $"\\b{Regex.Escape(replacement.Find)}\\b";
                    result = Regex.Replace(result, pattern, _ => replacement.Replace);
                }
                else
                {
                    result = result.Replace(replacement.Find, replacement.Replace, StringComparison.Ordinal);
                }
            }

            return result;
        }

        private string ApplyReplacement(string value)
        {
            if (replacementEntries == null || !replacementEntries.Any() || string.IsNullOrEmpty(value))
            {
                return value;
            }

            var result = value;

            foreach (var replacement in replacementEntries)
            {
                if (string.IsNullOrEmpty(replacement.Find))
                {
                    continue;
                }

                if (replacement.WholeWordMatch)
                {
                    var pattern = $"\\b{Regex.Escape(replacement.Find)}\\b";
                    result = Regex.Replace(result, pattern, _ => replacement.Replace);
                }
                else
                {
                    result = result.Replace(replacement.Find, replacement.Replace, StringComparison.Ordinal);
                }
            }

            return result;
        }

        private string DetermineCsvDelimiter(string? delimiter)
        {
            if (string.IsNullOrEmpty(delimiter))
            {
                return ";";
            }

            return delimiter.Length == 1 ? delimiter : delimiter[0].ToString();
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
