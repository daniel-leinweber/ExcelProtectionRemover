# ExcelProtectionRemover

CLI tool to remove the Worksheet Protection from an Excel file.

<!--toc:start-->
- [ExcelProtectionRemover](#excelprotectionremover)
  - [Usage](#usage)
    - [Build the Project from source](#build-the-project-from-source)
    - [Run the Tool](#run-the-tool)
  - [Build executable](#build-executable)
    - [Build the Project as Release](#build-the-project-as-release)
    - [Publish](#publish)
    - [Run the executable](#run-the-executable)
<!--toc:end-->

## Usage

### Build the Project from source

In the project directory, run:

```bash
dotnet build
```

### Run the Tool

Once built, run the tool from the command line with two arguments: the path to
the protected workbook and the desired output file path. For example:

```bash
dotnet run -- "ProtectedWorkbook.xlsx" "UnprotectedWorkbook.xlsx"
```

## Build executable

### Build the Project as Release

In the project directory, run:

```bash
dotnet build -c Release
```

### Publish

**Windows:**

```powershell
dotnet publish -r win-x64 -p:PublishSingleFile=true --self-contained true
```

**Linux:**

```bash
dotnet publish -r linux-x64 -p:PublishSingleFile=true --self-contained true
```

### Run the executable

**Windows:**

```powershell
ExcelProtectionRemover.exe <inputFile.xlsx> <outputFile.xlsx>
```

**Linux:**

```powershell
./ExcelProtectionRemover <inputFile.xlsx> <outputFile.xlsx>
```
