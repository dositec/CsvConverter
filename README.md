# CsvConverter

CsvConverter is a lightweight and flexible tool for converting CSV files to Excel (XLSX) format, with support for column mapping using a YAML configuration file. The application is built with .NET 8 and features a Windows Forms GUI.



## Features

- Convert CSV files to Excel (XLSX) format.
- Customize column mappings using a YAML configuration file.
- Automatically handle missing or mismatched columns with clear error logging.
- User-friendly GUI with support for selecting multiple files.
- Progress tracking and cancelation support during conversions.



## Installation

### Prerequisites
- **Windows 10 or later** (for running .NET 8 applications).
- **.NET 8 Runtime**: Required for the framework-dependent release. Download it from the [official .NET website](https://dotnet.microsoft.com/).

### Download
1. Visit the [Releases](https://github.com/KooleControls/CsvConverter/releases) page of this repository.
2. Choose between the two types of releases:
   - **Self-Contained**: Includes the .NET runtime. No additional installation is required. Recommended for most users.
   - **Framework-Dependent**: Requires the .NET runtime to be installed on your system. Smaller download size.
3. Download the appropriate zip file for your needs:
   - `CsvConverter-SelfContained-x.y.z.zip` (self-contained release).
   - `CsvConverter-FrameworkDependent-x.y.z.zip` (framework-dependent release).
4. Extract the downloaded archive to a folder of your choice.



## Usage

1. Launch the application by double-clicking the executable (`CsvConverter.exe`).
2. Add CSV files to convert by selecting **File > Add Files**.
3. Ensure the `Config.yaml` file is present in the application directory to define column mappings.
4. Click the **Convert** button to start processing the files.
5. Monitor progress using the built-in progress bar.
6. View detailed logs for any errors or missing columns in the log section.

### YAML Configuration Example

The `Config.yaml` file specifies the column mapping from the CSV input to the Excel output.

```yaml
Caption: "Report title"
CaptionBackgroundColor: "#4472C4"
HeaderBackgroundColor: "#D9E1F2"
ZebraColors:
  - "#F2F2F2"
  - "#FFFFFF"
Columns:
  - Input: "CSV Column Name 1"
    OutputIndex: 1
    Output: "Excel Header 1"
  - Input: "CSV Column Name 2"
    OutputIndex: 2
    Output: "Excel Header 2"
```

Place this file in the same directory as the application executable.



### Why Two Releases?

Providing both self-contained and framework-dependent releases ensures flexibility for users with different requirements:
- **Self-Contained**: For ease of use, especially on systems without .NET.
- **Framework-Dependent**: For reduced download size when .NET is already installed.

Choose the release that best suits your environment!



## Release Details

CsvConverter provides two types of releases to accommodate different user needs:

### Self-Contained Release
- **File Name**: `CsvConverter-SelfContained-x.y.z.zip`
- **Description**: Includes the .NET runtime, allowing the application to run without requiring the runtime to be installed on the system.
- **Use Case**: Recommended for most users, especially on systems without the .NET runtime installed.
- **Size**: Larger file size due to the included runtime.

### Framework-Dependent Release
- **File Name**: `CsvConverter-FrameworkDependent-x.y.z.zip`
- **Description**: Does not include the .NET runtime. Requires the .NET runtime to be installed on the system.
- **Use Case**: Ideal for users who already have the .NET runtime installed and want a smaller download size.
- **Size**: Smaller file size since it relies on the installed runtime.



## Development

### Prerequisites
- Visual Studio 2022 or later with .NET Desktop Development workload installed.

### Building the Project
1. Clone this repository:
   ```bash
   git clone https://github.com/KooleControls/CsvConverter/csvconverter.git
   cd csvconverter
   ```
2. Open the project in Visual Studio.
3. Build the solution to generate the executable.



## Contributing

Contributions are welcome! Please follow these steps:
1. Fork the repository.
2. Create a new branch for your feature or bug fix.
3. Commit and push your changes.
4. Submit a pull request.



## License

This project is licensed under the [MIT License](LICENSE).



## Acknowledgments

- [ClosedXML](https://github.com/ClosedXML/ClosedXML) for working with Excel files.
- [YamlDotNet](https://github.com/aaubry/YamlDotNet) for parsing YAML configuration files.

