using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using YamlDotNet.Core.Tokens;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;

namespace CsvConverter
{

    public partial class Form1 : Form
    {

        private string yamlFilePath;
        private readonly BindingList<WorkItem> workItems;
        private CsvColumnMapper janConverter;
        private readonly RichTextBoxLoggerFactory loggerFactory;
        private readonly ILogger<Form1> logger;
        private readonly string settingsFilePath;

        public Form1()
        {
            InitializeComponent();

            var version = Assembly.GetExecutingAssembly().GetName().Version?.ToString() ?? "Unknown";
            Text = $"Csv converter - Version {version}";

            settingsFilePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "CsvConverter", "settings.txt");
            Directory.CreateDirectory(Path.GetDirectoryName(settingsFilePath)!);

            yamlFilePath = LoadConfigPath() ?? "Config.yaml";

            workItems = new BindingList<WorkItem>();
            loggerFactory = new RichTextBoxLoggerFactory(richTextBox_Log);
            logger = loggerFactory.CreateLogger<Form1>();
            janConverter = new CsvColumnMapper(loggerFactory.CreateLogger<CsvColumnMapper>(), yamlFilePath);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // Load the column mapping from the YAML file
            toolStrip1.AddMenuItem("File/Clear").WithAction(workItems.Clear);
            toolStrip1.AddMenuItem("File/Add files").WithAction(SelectFiles);
            toolStrip1.AddMenuItem("File/Select Config").WithAction(SelectConfig);
            toolStrip1.AddMenuItem("File/Exit").WithAction(Close);

            UpdateTitleWithConfigPath();

            ConfigureListView();
        }

        private void ConfigureListView()
        {
            // Set the view to details to show headers and columns
            listView1.View = View.Details;

            // Add columns for headers
            listView1.Columns.Add("Input File", 300);   // Column for Input File
            listView1.Columns.Add("Output File", 300);  // Column for Output File
            listView1.Columns.Add("Status", 100);       // Column for Status

            // Enable virtual mode
            listView1.VirtualMode = true;
            listView1.VirtualListSize = workItems.Count;

            // Handle the RetrieveVirtualItem event to populate rows
            listView1.RetrieveVirtualItem += (s, e) =>
            {
                if (e.ItemIndex >= 0 && e.ItemIndex < workItems.Count)
                {
                    var item = workItems[e.ItemIndex];
                    e.Item = new ListViewItem(new[]
                    {
                item.InputFile,
                item.OuputFile,
                item.Status.ToString()
            });
                }
            };

            // Update the virtual list size when the data changes
            workItems.ListChanged += (s, e) =>
            {
                listView1.VirtualListSize = workItems.Count;
                listView1.Refresh();
            };
        }


        private void SelectConfig()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "YAML Files (*.yaml;*.yml)|*.yaml;*.yml|All Files (*.*)|*.*",
                Title = "Select Config File",
                FileName = yamlFilePath
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                yamlFilePath = openFileDialog.FileName;
                SaveConfigPath(yamlFilePath);
                janConverter = new CsvColumnMapper(loggerFactory.CreateLogger<CsvColumnMapper>(), yamlFilePath);
                UpdateTitleWithConfigPath();
                logger.LogInformation($"Config file changed to: {yamlFilePath}");
            }
        }

        private void UpdateTitleWithConfigPath()
        {
            var version = Assembly.GetExecutingAssembly().GetName().Version?.ToString() ?? "Unknown";
            Text = $"Csv converter - Version {version} - Config: {yamlFilePath}";
        }

        private string? LoadConfigPath()
        {
            if (File.Exists(settingsFilePath))
            {
                try
                {
                    return File.ReadAllText(settingsFilePath).Trim();
                }
                catch
                {
                    return null;
                }
            }
            return null;
        }

        private void SaveConfigPath(string path)
        {
            try
            {
                File.WriteAllText(settingsFilePath, path);
            }
            catch (Exception ex)
            {
                logger.LogWarning($"Failed to save config path: {ex.Message}");
            }
        }

        private void SelectFiles()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Multiselect = true, // Allow selection of multiple files
                Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*", // Filter to show only CSV files by default
                Title = "Select CSV Files"
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                // The selected file paths are stored in the FileNames property
                string[] selectedFiles = openFileDialog.FileNames;

                // Process the selected files
                foreach (string file in selectedFiles)
                {
                    // Check if the file is already in the list
                    if (workItems.Any(w => string.Equals(w.InputFile, file, StringComparison.OrdinalIgnoreCase)))
                    {
                        logger.LogWarning($"File '{file}' is already in the list and will be skipped.");
                        continue;
                    }

                    WorkItem item = new()
                    {
                        InputFile = file,
                        OuputFile = Path.ChangeExtension(file, ".xlsx"), // Change extension to .xlsx
                        Status = WorkItemStatus.Todo
                    };
                    workItems.Add(item);
                }
            }
        }

        

        CancellationTokenSource cancellationTokenSource = new CancellationTokenSource();
        private async void button_Convert_Click(object sender, EventArgs e)
        {
            button_Convert.Enabled = false;
            toolStrip1.Enabled = false;
            progressBar1.Value = 0;
            progressBar1.Maximum = workItems.Count; // Set progress bar maximum to the number of work items
            cancellationTokenSource = new CancellationTokenSource();

            try
            {
                int completedCount = 0;

                foreach (var item in workItems)
                {
                    if (cancellationTokenSource.Token.IsCancellationRequested)
                        break;

                    await HandleWorkItem(item, cancellationTokenSource.Token);

                    completedCount++;
                    progressBar1.Value = completedCount; // Update progress bar value

                    // Log progress to the logger
                    logger.LogInformation($"Progress: {completedCount}/{workItems.Count}");
                    listView1.Refresh();
                }
            }
            catch (OperationCanceledException)
            {
                logger.LogInformation("Conversion was canceled by the user.");
            }
            catch (Exception ex)
            {
                logger.LogError($"An unexpected error occurred: {ex.Message}");
            }
            finally
            {
                button_Convert.Enabled = true;
                toolStrip1.Enabled = true;
            }
        }

        private async Task HandleWorkItem(WorkItem item, CancellationToken token)
        {
            try
            {
                item.Status = WorkItemStatus.Todo; // Start as TODO

                bool success = await janConverter.ConvertAsync(item.InputFile, item.OuputFile, token);

                if (success)
                {
                    item.Status = WorkItemStatus.Done;
                }
                else
                {
                    item.Status = WorkItemStatus.Error;
                }
            }
            catch (OperationCanceledException)
            {
                item.Status = WorkItemStatus.Todo; // Revert status if canceled
                throw;
            }
            catch (Exception ex)
            {
                item.Status = WorkItemStatus.Error; // Mark as ERROR on failure
                logger.LogError($"Error converting file {item.InputFile}: {ex.Message}");
            }
        }

        private void button_Cancel_Click(object sender, EventArgs e)
        {
            cancellationTokenSource.Cancel();
        }
    }
}
