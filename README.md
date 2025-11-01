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

## Development

For technical details about the architecture, design decisions, and implementation, see [ARCHITECTURE.md](ARCHITECTURE.md).

## License

This project is provided as-is for text extraction from Excel files.
