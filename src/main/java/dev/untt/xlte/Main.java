package dev.untt.xlte;

import picocli.CommandLine;
import picocli.CommandLine.Command;
import picocli.CommandLine.Parameters;
import picocli.CommandLine.Option;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.Callable;
import java.util.stream.Stream;

@Command(
    name = "xlte",
    version = "xlte 1.0.0",
    description = "Extract text from Excel (.xlsx) files"
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
        names = {"-q", "--quiet"},
        description = "Suppress file headers when processing multiple files"
    )
    private boolean quiet;

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
        var exitCode = new CommandLine(new Main()).execute(args);
        System.exit(exitCode);
    }

    @Override
    public Integer call() {
        // Validate that exactly one of -f or -d is specified
        if (file == null && directory == null) {
            System.err.println("Error: Either --file or --dir must be specified");
            System.err.println("Use --help for usage information");
            return 1;
        }

        if (file != null && directory != null) {
            System.err.println("Error: Cannot specify both --file and --dir");
            return 1;
        }

        try {
            // Detect output mode: terminal vs redirected/piped
            var isTerminal = System.console() != null;
            var outputMode = isTerminal ?
                ExcelTextExtractor.OutputMode.TERMINAL :
                ExcelTextExtractor.OutputMode.TSV;

            var extractor = new ExcelTextExtractor(outputMode);
            var filesToProcess = new ArrayList<File>();

            if (file != null) {
                // Single file mode
                if (!file.exists()) {
                    System.err.println("Error: File not found: " + file);
                    return 1;
                }

                if (!file.isFile()) {
                    System.err.println("Error: Not a file: " + file);
                    return 1;
                }

                if (!file.getName().toLowerCase().endsWith(".xlsx")) {
                    System.err.println("Error: Only .xlsx files are supported");
                    return 1;
                }

                filesToProcess.add(file);
            } else {
                // Directory mode
                if (!directory.exists()) {
                    System.err.println("Error: Directory not found: " + directory);
                    return 1;
                }

                if (!directory.isDirectory()) {
                    System.err.println("Error: Not a directory: " + directory);
                    return 1;
                }

                // Find all .xlsx files in directory
                filesToProcess.addAll(findExcelFiles(directory.toPath(), recursive));

                if (filesToProcess.isEmpty()) {
                    System.err.println("No .xlsx files found in directory: " + directory);
                    return 1;
                }
            }

            // Process all files
            var showFileName = !quiet && filesToProcess.size() > 1;
            for (var excelFile : filesToProcess) {
                processFile(extractor, excelFile, showFileName);
            }

            return 0;
        } catch (Exception e) {
            System.err.println("Error: " + e.getMessage());
            if (System.getenv("DEBUG") != null) {
                e.printStackTrace();
            }
            return 1;
        }
    }

    private List<File> findExcelFiles(Path directory, boolean recursive) throws IOException {
        var excelFiles = new ArrayList<File>();

        try (Stream<Path> paths = recursive ? Files.walk(directory) : Files.list(directory)) {
            paths.filter(Files::isRegularFile)
                 .filter(p -> p.toString().toLowerCase().endsWith(".xlsx"))
                 .sorted()
                 .forEach(p -> excelFiles.add(p.toFile()));
        }

        return excelFiles;
    }

    private void processFile(ExcelTextExtractor extractor, File file, boolean showFileName) {
        try {
            var text = extractor.extractText(file);
            if (!text.isEmpty()) {
                System.out.println(text);
            }
        } catch (Exception e) {
            System.err.println("Error processing file " + file.getPath() + ": " + e.getMessage());
        }
    }
}
