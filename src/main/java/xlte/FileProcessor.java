package xlte;

import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;
import java.util.stream.Collectors;

/**
 * Handles file I/O and delegates to WorkbookExtractor for content extraction.
 * This class is responsible for opening Excel files and managing resources.
 * The actual extraction logic is delegated to WorkbookExtractor, which operates
 * purely on POI Workbook objects without any file I/O dependencies.
 */
public class FileProcessor {
    private final WorkbookExtractor workbookExtractor;

    /**
     * Creates a new FileProcessor with a default WorkbookExtractor.
     */
    public FileProcessor() {
        this.workbookExtractor = new WorkbookExtractor();
    }

    /**
     * Extracts all text content from an Excel file.
     *
     * @param file The Excel file to process
     * @return An unmodifiable list of extracted items (cells and shapes)
     * @throws IOException If the file cannot be read
     */
    public List<ExtractedItem> processFile(File file) throws IOException {
        var filePath = file.getPath();

        try (var fis = new FileInputStream(file);
             var workbook = WorkbookFactory.create(fis)) {

            return workbookExtractor
                .extract(workbook, filePath)
                .collect(Collectors.toUnmodifiableList());
        }
    }
}
