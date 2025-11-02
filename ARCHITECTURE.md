# Architecture

This document describes the technical architecture and design decisions of xlte.

## Overview

The project follows a functional programming approach with clear separation of concerns. The codebase emphasizes immutability, pure functions, and type safety using modern Java features.

## Three-Layer Architecture

The application is organized into three distinct layers, each with a specific responsibility:

### 1. File I/O Layer (`FileProcessor`)

- **Responsibility**: Handles file system operations and resource management
- **Key Features**:
  - Manages file input streams with try-with-resources
  - Creates POI Workbook instances from files
  - Handles IOException and resource cleanup
  - Delegates actual extraction to WorkbookExtractor
- **Dependencies**: Java I/O, Apache POI WorkbookFactory
- **Testability**: Can be tested with actual files or mocked file systems

### 2. Workbook Processing Layer (`WorkbookExtractor`)

- **Responsibility**: Processes Apache POI Workbook objects without file I/O concerns
- **Key Features**:
  - Iterates over all sheets in a workbook
  - Aggregates results from multiple sheets
  - Testable without file system access (can pass in-memory Workbook)
  - Delegates sheet-level extraction to SheetExtractor
- **Dependencies**: Apache POI Workbook API, SheetExtractor
- **Testability**: Highly testable - can create Workbook objects in memory for unit tests

### 3. Sheet Extraction Layer (`SheetExtractor`)

- **Responsibility**: Pure POI logic for extracting cells and shapes from individual sheets
- **Key Features**:
  - Extracts text from all cells in a sheet
  - Extracts text from shapes (text boxes, etc.)
  - Uses Stream API for functional data processing
  - No side effects or external dependencies
  - Pure functions throughout
- **Dependencies**: Apache POI Sheet/Row/Cell APIs, DataFormatter
- **Testability**: Easiest layer to test - pure functions with no side effects

### Benefits of This Architecture

1. **Separation of Concerns**: Each layer has a single, well-defined responsibility
2. **Testability**: Layers can be tested independently without file I/O overhead
3. **Reusability**: WorkbookExtractor can be used with in-memory workbooks
4. **Maintainability**: Changes to file handling don't affect extraction logic
5. **Functional Purity**: Core extraction logic (SheetExtractor) is purely functional

## Functional Programming Features

This project heavily emphasizes functional programming principles:

### Pure Functions

Most methods in the codebase are pure functions:
- **Referential Transparency**: Same inputs always produce same outputs
- **No Side Effects**: Functions don't modify external state
- **Testability**: Pure functions are easier to test and reason about

Examples:
- `SheetExtractor.extractCells()` - pure transformation of Sheet to Stream of items
- `Result.map()` and `Result.flatMap()` - pure monadic operations
- All validation methods in `Main.java` - pure predicates

### Immutability

- **Records**: All data models are immutable records (`CellItem`, `ShapeItem`, `Result.Success`, `Result.Failure`, `InputMode.SingleFile`, `InputMode.Directory`)
- **Immutable Collections**: Using `Collectors.toUnmodifiableList()` throughout
- **No Mutable State**: Formatters are stateless (no instance fields)

### Stream API

Declarative data processing using Java Streams:
- `map()` - Transform elements
- `flatMap()` - Flatten nested streams
- `filter()` - Select elements
- `collect()` - Aggregate results

Example from `SheetExtractor`:
```java
return StreamSupport.stream(sheet.spliterator(), false)
    .flatMap(row -> StreamSupport.stream(row.spliterator(), false))
    .filter(cell -> cell != null)
    .map(cell -> extractCellValue(cell, dataFormatter))
    .filter(cv -> !cv.isEmpty())
    .map(cv -> new CellItem(filePath, sheetName, cv.address(), cv.value()));
```

### Result Monad

Type-safe error handling without exceptions for expected error cases:

```java
public Integer call() {
    var result = validateInput()
        .flatMap(this::resolveFiles)
        .map(this::processAllFilesAndGetExitCode);

    result.ifFailure(System.err::println);
    return result.orElse(1);
}
```

Benefits:
- Explicit error handling in type signatures
- Compose error-prone operations with `flatMap()`
- Separation of error handling from business logic
- No try-catch blocks for expected errors

### Sealed Interfaces

Type-safe discriminated unions for domain models:

- **`ExtractedItem`**: Either a `CellItem` or `ShapeItem`
- **`InputMode`**: Either `SingleFile` or `Directory`
- **`Result<T>`**: Either `Success<T>` or `Failure<T>`

Benefits:
- Exhaustive pattern matching in switch expressions
- Compiler ensures all cases are handled
- Cannot add new implementations outside the defining file

### Stateless Components

All formatters (`TerminalFormatter`, `TsvFormatter`) are stateless:
- No mutable instance fields
- Thread-safe by design
- Can be reused across multiple invocations
- Easier to test and reason about

## Java 25 Features Used

This project takes advantage of modern Java features:

### Records (Java 14+)

Immutable data carriers with automatic implementations of:
- Constructor
- Getters
- `equals()` and `hashCode()`
- `toString()`

Used for:
- `CellItem(String filePath, String sheetName, String cellAddress, String content)`
- `ShapeItem(String filePath, String sheetName, String content)`
- `InputMode.SingleFile(File file)`
- `InputMode.Directory(File dir, boolean recursive)`
- `Result.Success<T>(T value)`
- `Result.Failure<T>(String error)`

### Sealed Interfaces (Java 17+)

Restrict which classes can implement an interface:

```java
public sealed interface Result<T> permits Result.Success, Result.Failure {
    record Success<T>(T value) implements Result<T> {}
    record Failure<T>(String error) implements Result<T> {}
}
```

### Pattern Matching for Switch (Java 21+)

Type patterns in switch expressions:

```java
return switch (result) {
    case Result.Success<String> success -> success.value();
    case Result.Failure<String> failure -> {
        System.err.println(failure.error());
        yield "";
    }
};
```

### var Keyword (Java 10+)

Local variable type inference for cleaner code:

```java
var formatter = new TerminalFormatter();
var items = fileProcessor.processFile(file);
var output = formatter.format(items);
```

### Stream API (Java 8+)

Functional data processing with streams - used extensively throughout the codebase.

### try-with-resources (Java 7+, enhanced in Java 9+)

Automatic resource management:

```java
try (var fis = new FileInputStream(file);
     var workbook = WorkbookFactory.create(fis)) {
    return workbookExtractor.extract(workbook, filePath)
        .collect(Collectors.toUnmodifiableList());
}
```

## Code Structure

```
src/main/java/xlte/
├── Main.java                 - CLI entry point with picocli
├── FileProcessor.java        - File I/O layer
├── WorkbookExtractor.java    - Workbook processing layer
├── SheetExtractor.java       - Sheet extraction layer (pure POI logic)
├── Result.java               - Result monad for error handling
├── InputMode.java            - Input mode sealed interface
├── ExtractedItem.java        - Sealed interface for extracted items
├── CellItem.java             - Record for cell data
├── ShapeItem.java            - Record for shape data
├── OutputFormatter.java      - Formatter interface
├── TerminalFormatter.java    - Human-readable output (stateless)
└── TsvFormatter.java         - Machine-readable TSV output (stateless)
```

### Component Relationships

```
Main (CLI)
  ├─> validates input using Result monad
  ├─> FileProcessor
  │     └─> WorkbookExtractor
  │           └─> SheetExtractor (pure POI logic)
  └─> OutputFormatter (TerminalFormatter or TsvFormatter)
```

## Design Decisions

### Why Functional Programming?

**Advantages:**
1. **Testability**: Pure functions are easier to test - no setup/teardown needed
2. **Maintainability**: Immutability eliminates whole classes of bugs
3. **Composability**: Small pure functions can be composed into larger operations
4. **Concurrency**: Immutable data structures are naturally thread-safe
5. **Reasoning**: Easier to understand code without hidden side effects

**Trade-offs:**
- Slightly more verbose in some cases (creating new objects instead of mutation)
- Learning curve for developers unfamiliar with functional concepts
- Overall: Benefits far outweigh the costs for this use case

### Why Three Layers?

**Separation of file I/O from business logic:**
1. **Testing**: Can test extraction logic without creating files
2. **Reusability**: WorkbookExtractor can process in-memory workbooks
3. **Performance**: Can profile each layer independently
4. **Flexibility**: Easy to add new input sources (network, database, etc.)

**Separation of sheet-level logic:**
1. **Single Responsibility**: Each extractor has one clear purpose
2. **Purity**: Sheet extraction is purely functional with no I/O
3. **Testing**: Easy to create test sheets for unit tests

### Error Handling with Result Monad

**Why not exceptions?**
- Exceptions for expected errors (file not found, invalid input) hide control flow
- Try-catch blocks mix error handling with business logic
- Exceptions are not visible in method signatures

**Why Result monad?**
1. **Explicit**: Error handling is explicit in type signatures
2. **Composable**: Can chain operations with `flatMap()` and stop on first error
3. **Type-safe**: Compiler ensures errors are handled
4. **Functional**: Fits naturally with functional programming style

**When to use exceptions vs Result:**
- **Result**: Expected errors (validation, file not found, no Excel files in directory)
- **Exceptions**: Unexpected errors (IOException during read, POI parsing errors)

### Output Format Auto-Detection

The tool automatically switches between Terminal and TSV mode based on output destination:
- **Terminal Mode**: Human-readable format when outputting directly to console
- **TSV Mode**: Machine-readable format when piped or redirected

**Implementation:**
```java
var isTerminal = System.console() != null;
return isTerminal ? new TerminalFormatter() : new TsvFormatter();
```

**Benefits:**
- Best user experience for both interactive and scripted use
- No need for explicit flags
- Follows Unix philosophy (similar to `ls`, `ripgrep`, etc.)

## Future Enhancements

Potential areas for improvement:

1. **Unit Tests**: Add comprehensive unit tests (JUnit 5)
   - Test each layer independently
   - Test Result monad behavior
   - Test formatters with various inputs

2. **Parallel Processing**: Use parallel streams for large directories
   - `Files.walk(directory).parallel()`
   - Need to ensure thread-safety (already achieved via immutability)

3. **Streaming Large Files**: For very large Excel files, process incrementally
   - Currently loads entire file into memory
   - Could stream sheet-by-sheet

4. **Additional Output Formats**: JSON, CSV, etc.
   - Easy to add new formatters implementing `OutputFormatter`

5. **Filter Options**: Allow filtering by sheet name, cell range, content pattern
   - Would fit naturally into Stream pipeline

6. **Progress Reporting**: Show progress for large operations
   - Need to balance with functional purity
   - Could use side-effect isolated component
