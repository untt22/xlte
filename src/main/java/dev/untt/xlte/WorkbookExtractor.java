package dev.untt.xlte;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.stream.IntStream;
import java.util.stream.Stream;

/**
 * Extracts text content from Apache POI Workbook objects.
 * This class has no file I/O dependencies and works purely with POI objects.
 * It can be tested without any file operations by passing Workbook instances directly.
 */
public class WorkbookExtractor {

    /**
     * Extracts all text content from a Workbook.
     *
     * @param workbook The POI Workbook to extract from
     * @param filePath The file path to associate with extracted items
     * @return A stream of extracted items (cells and shapes)
     */
    public Stream<ExtractedItem> extract(Workbook workbook, String filePath) {
        return IntStream.range(0, workbook.getNumberOfSheets())
            .mapToObj(workbook::getSheetAt)
            .flatMap(sheet -> extractFromSheet(sheet, filePath));
    }

    /**
     * Extracts content from a single sheet.
     *
     * @param sheet The sheet to extract from
     * @param filePath The file path to associate with extracted items
     * @return A stream of extracted items
     */
    private Stream<ExtractedItem> extractFromSheet(Sheet sheet, String filePath) {
        var sheetName = sheet.getSheetName();
        var sheetExtractor = new SheetExtractor(filePath, sheetName);

        var cellStream = sheetExtractor.extractCells(sheet);
        var shapeStream = sheetExtractor.extractShapes(sheet);

        return Stream.concat(cellStream, shapeStream);
    }
}
