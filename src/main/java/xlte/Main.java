package xlte;

import picocli.CommandLine;
import picocli.CommandLine.Command;
import picocli.CommandLine.Option;

import java.io.File;
import java.io.IOException;
import java.io.PrintStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.Stream;
import java.util.concurrent.Callable;

@Command(
    name = "xlte",
    version = "xlte 1.0.0",
    description = "Extract text from Excel (.xls, .xlsx, .xlsm) files"
)
public class Main implements Callable<Integer> {

    @Option(
        names = {"-f", "--file"},
        description = "Path to an Excel file to process",
        paramLabel = "FILE"
    )
    private File file;

    @Option(
        names = {"-d", "--dir"},
        description = "Path to a directory containing Excel files",
        paramLabel = "DIRECTORY"
    )
    private File directory;

    @Option(
        names = {"-r", "--recursive"},
        description = "Recursively process directories (default: true)",
        defaultValue = "true"
    )
    private boolean recursive;

    @Option(
        names = {"-h", "--help"},
        usageHelp = true,
        description = "Show this help message and exit"
    )
    private boolean helpRequested;

    @Option(
        names = {"-v", "--version"},
        versionHelp = true,
        description = "Print version information and exit"
    )
    private boolean versionRequested;

    public static void main(String[] args) {
        // Set UTF-8 encoding for stdout to ensure proper output encoding
        try {
            System.setOut(new PrintStream(System.out, true, StandardCharsets.UTF_8));
        } catch (Exception e) {
            // If setting encoding fails, continue with default
        }

        var exitCode = new CommandLine(new Main()).execute(args);
        System.exit(exitCode);
    }

    @Override
    public Integer call() {
        var result = validateInput()
            .flatMap(this::resolveFiles)
            .map(this::processAllFilesAndGetExitCode);

        result.ifFailure(System.err::println);
        return result.orElse(1);
    }

    /**
     * Validates that exactly one of --file or --dir is specified.
     * Pure function with no side effects.
     */
    private Result<InputMode> validateInput() {
        if (file == null && directory == null) {
            return new Result.Failure<>(
                "Error: Either --file or --dir must be specified\nUse --help for usage information"
            );
        }

        if (file != null && directory != null) {
            return new Result.Failure<>(
                "Error: Cannot specify both --file and --dir"
            );
        }

        InputMode mode = file != null
            ? new InputMode.SingleFile(file)
            : new InputMode.Directory(directory, recursive);

        return new Result.Success<>(mode);
    }

    /**
     * Resolves the input mode to a list of files to process.
     * Pure function except for file system access.
     */
    private Result<List<File>> resolveFiles(InputMode mode) {
        return switch (mode) {
            case InputMode.SingleFile(var f) -> validateFile(f).map(List::of);
            case InputMode.Directory(var dir, var recursive) ->
                validateDirectory(dir).flatMap(d -> findExcelFilesResult(d.toPath(), recursive));
        };
    }

    /**
     * Validates a single file.
     * Pure function except for file system access.
     */
    private Result<File> validateFile(File file) {
        return checkFileExists(file)
            .flatMap(this::checkIsRegularFile)
            .flatMap(this::checkHasExcelExtension);
    }

    private Result<File> checkFileExists(File file) {
        return file.exists()
            ? new Result.Success<>(file)
            : new Result.Failure<>("Error: File not found: " + file);
    }

    private Result<File> checkIsRegularFile(File file) {
        return file.isFile()
            ? new Result.Success<>(file)
            : new Result.Failure<>("Error: Not a file: " + file);
    }

    private Result<File> checkHasExcelExtension(File file) {
        var name = file.getName().toLowerCase();
        boolean isExcel = name.endsWith(".xls") ||
                         name.endsWith(".xlsx") ||
                         name.endsWith(".xlsm");
        return isExcel
            ? new Result.Success<>(file)
            : new Result.Failure<>("Error: Only .xls, .xlsx, and .xlsm files are supported");
    }

    /**
     * Validates a directory.
     * Pure function except for file system access.
     */
    private Result<File> validateDirectory(File dir) {
        if (!dir.exists()) {
            return new Result.Failure<>("Error: Directory not found: " + dir);
        }
        if (!dir.isDirectory()) {
            return new Result.Failure<>("Error: Not a directory: " + dir);
        }
        return new Result.Success<>(dir);
    }

    /**
     * Finds Excel files in a directory and wraps the result in a Result type.
     */
    private Result<List<File>> findExcelFilesResult(Path directory, boolean recursive) {
        try {
            var files = findExcelFiles(directory, recursive);
            if (files.isEmpty()) {
                return new Result.Failure<>("No Excel files found in directory: " + directory);
            }
            return new Result.Success<>(files);
        } catch (IOException e) {
            return new Result.Failure<>("Error reading directory: " + e.getMessage());
        }
    }

    /**
     * Finds all Excel files in a directory.
     * Functional implementation using Stream.collect().
     */
    private List<File> findExcelFiles(Path directory, boolean recursive) throws IOException {
        try (Stream<Path> paths = recursive ? Files.walk(directory) : Files.list(directory)) {
            return paths
                .filter(Files::isRegularFile)
                .filter(this::isExcelFile)
                .sorted()
                .map(Path::toFile)
                .collect(Collectors.toUnmodifiableList());
        }
    }

    /**
     * Checks if a path represents an Excel file.
     * Pure function.
     */
    private boolean isExcelFile(Path path) {
        var name = path.toString().toLowerCase();
        return name.endsWith(".xls") ||
               name.endsWith(".xlsx") ||
               name.endsWith(".xlsm");
    }

    /**
     * Processes all files and returns exit code.
     * This is where side effects (printing output) occur.
     */
    private int processAllFilesAndGetExitCode(List<File> files) {
        var formatter = createFormatter();
        var fileProcessor = new FileProcessor();

        // Process files and collect results
        files.stream()
            .map(file -> processFile(fileProcessor, formatter, file))
            .forEach(this::handleResult);

        return 0;
    }

    /**
     * Creates the appropriate formatter based on output destination.
     */
    private OutputFormatter createFormatter() {
        var isTerminal = System.console() != null;
        return isTerminal ? new TerminalFormatter() : new TsvFormatter();
    }

    /**
     * Processes a single file and returns the result.
     * Pure function except for file I/O.
     */
    private Result<String> processFile(FileProcessor fileProcessor, OutputFormatter formatter, File file) {
        try {
            var items = fileProcessor.processFile(file);
            var output = formatter.format(items);
            return new Result.Success<>(output);
        } catch (Exception e) {
            return new Result.Failure<>(
                "Error processing file " + file.getPath() + ": " + e.getMessage()
            );
        }
    }

    /**
     * Handles the result of processing a file.
     * This is where side effects (printing) occur.
     */
    private void handleResult(Result<String> result) {
        switch (result) {
            case Result.Success<String> success -> {
                if (!success.value().isEmpty()) {
                    System.out.print(success.value());
                }
            }
            case Result.Failure<String> failure -> {
                System.err.println(failure.error());
            }
        }
    }
}
