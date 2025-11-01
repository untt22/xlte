# xlte

Excel Text Extractor - Extract text from Excel (.xls, .xlsx, .xlsm) files

## Features

- Extract text from all cells in Excel workbooks
- Extract text from shapes (text boxes, etc.)
- Support for multiple sheets
- Support for multiple Excel formats:
  - `.xls` (Excel 97-2003)
  - `.xlsx` (Excel 2007+)
  - `.xlsm` (Excel 2007+ with macros)
- Command-line interface

## Requirements

- Java 25 or higher
- Maven 3.6 or higher

## Build

```bash
mvn clean package
```

This will create an executable JAR file with all dependencies in `target/xlte-1.0.0.jar`

## Usage

```bash
java -jar target/xlte-1.0.0.jar [OPTIONS]
```

### Options

- `-f, --file FILE` - Path to an Excel file to process
- `-d, --dir DIRECTORY` - Path to a directory containing Excel files
- `-r, --recursive` - Recursively process directories (default: true)
- `-h, --help` - Show help message and exit
- `-v, --version` - Print version information and exit

**Note:** Either `--file` or `--dir` must be specified, but not both.

### Examples

**Extract text from a single file:**
```bash
java -jar target/xlte-1.0.0.jar --file sample.xlsx
# or short form:
java -jar target/xlte-1.0.0.jar -f sample.xlsx
```

**Extract text from all Excel files in a directory (recursive):**
```bash
java -jar target/xlte-1.0.0.jar --dir ./samples
# or short form:
java -jar target/xlte-1.0.0.jar -d ./samples
```

**Process only files in the specified directory (non-recursive):**
```bash
java -jar target/xlte-1.0.0.jar --dir ./samples --recursive=false
```

**Show help:**
```bash
java -jar target/xlte-1.0.0.jar --help
```

## Output Format

The tool automatically adapts its output format based on whether the output is going to a terminal or being redirected/piped (similar to ripgrep):

### Terminal Mode (Human-Readable)

When output goes directly to a terminal, the format is designed for human readability:

```
=== samples/sample.xlsx ===
[Sheet1:A1] Cell1Content
[Sheet1:A2] Cell2Content
[Sheet1:A3] Cell3:WithColon
[Sheet1:Shape] ShapeText1

=== samples/data.xlsx ===
[Sheet2:C5] AnotherCell
```

**Characteristics:**
- File headers with separators
- `[Sheet:Cell]` prefix for easy visual scanning
- Newlines preserved in their original form
- Blank line between files

### TSV Mode (Machine-Readable)

When output is redirected to a file or piped to another command, the format switches to tab-separated values:

```
samples/sample.xlsx	Sheet1	A1	Cell1Content
samples/sample.xlsx	Sheet1	A2	Cell2Content
samples/sample.xlsx	Sheet1	A3	Cell3:WithColon
samples/sample.xlsx	Sheet1	Shape	ShapeText1
samples/data.xlsx	Sheet2	C5	AnotherCell
```

**Columns:**
1. **File Path** - Path to the Excel file
2. **Sheet Name** - Name of the sheet
3. **Cell Position** - Cell address (e.g., A1, B2) or "Shape" for text boxes
4. **Text Content** - The extracted text

**Characteristics:**
- Tab-separated columns for easy parsing
- Newlines escaped as `\n`
- Compatible with grep, awk, and spreadsheet software

### Usage Examples

**View in terminal (human-readable):**
```bash
java -jar target/xlte-1.0.0.jar -d ./data
```

**Save to file and search (TSV):**
```bash
java -jar target/xlte-1.0.0.jar -d ./data > index.txt
grep "keyword" index.txt
```

**Pipe to grep (TSV):**
```bash
java -jar target/xlte-1.0.0.jar -d ./data | grep "keyword"
```

**Import into spreadsheet (TSV):**
```bash
java -jar target/xlte-1.0.0.jar -d ./data > output.tsv
# Open output.tsv in Excel or other spreadsheet software
```

## Dependencies

- Apache POI 5.4.1 - For Excel file processing
- picocli 4.7.7 - For command-line interface

## Architecture

The project follows a functional programming approach with clear separation of concerns:

### Three-Layer Architecture

1. **File I/O Layer** (`FileProcessor`)
   - Handles file system operations
   - Manages resource lifecycle (try-with-resources)
   - Delegates to WorkbookExtractor

2. **Workbook Processing Layer** (`WorkbookExtractor`)
   - Processes POI Workbook objects
   - Testable without file I/O
   - Delegates sheet processing to SheetExtractor

3. **Sheet Extraction Layer** (`SheetExtractor`)
   - Pure POI logic for extracting cells and shapes
   - Uses Stream API for functional data processing
   - No side effects or external dependencies

### Functional Programming Features

- **Pure Functions**: Most methods have no side effects and are referentially transparent
- **Immutability**: Records and immutable collections throughout
- **Stream API**: Declarative data processing with `map()`, `flatMap()`, `filter()`, `collect()`
- **Result Monad**: Type-safe error handling without exceptions (see `Result.java`)
- **Sealed Interfaces**: Type-safe discriminated unions for domain models
- **Stateless Components**: Thread-safe formatters with no mutable state

## Java 25 Features Used

This project takes advantage of the latest Java features:

- **Records**: Immutable data carriers for `CellItem`, `ShapeItem`, `InputMode`, `Result.Success`, `Result.Failure`
- **Sealed interfaces**: `ExtractedItem`, `InputMode`, and `Result` sealed interfaces restrict permitted implementations
- **Pattern matching for switch**: Type patterns in switch expressions throughout
- **var keyword**: Local variable type inference for cleaner code
- **Stream API**: Functional data processing with streams
- **try-with-resources**: Automatic resource management
