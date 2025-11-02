package xlte;

import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFSimpleShape;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;

import java.util.Optional;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;

/**
 * Extracts text from Excel sheets (cells and shapes).
 * This class contains the core extraction logic for POI Sheet objects.
 * It operates purely on POI objects without any file I/O dependencies.
 */
public class SheetExtractor {
    private final String filePath;
    private final String sheetName;

    /**
     * Creates a new SheetExtractor.
     *
     * @param filePath The file path to associate with extracted items
     * @param sheetName The sheet name to associate with extracted items
     */
    public SheetExtractor(String filePath, String sheetName) {
        this.filePath = filePath;
        this.sheetName = sheetName;
    }

    /**
     * Extracts all non-empty cells from the sheet.
     *
     * @param sheet The sheet to extract cells from
     * @return A stream of extracted cell items
     */
    public Stream<ExtractedItem> extractCells(Sheet sheet) {
        var dataFormatter = new DataFormatter();

        return StreamSupport.stream(sheet.spliterator(), false)
            .flatMap(row -> StreamSupport.stream(row.spliterator(), false))
            .filter(cell -> cell != null)
            .map(cell -> extractCellValue(cell, dataFormatter))
            .filter(cv -> !cv.isEmpty())
            .map(cv -> new CellItem(filePath, sheetName, cv.address(), cv.value()));
    }

    /**
     * Extracts all text from shapes in the sheet.
     *
     * @param sheet The sheet to extract shapes from
     * @return A stream of extracted shape items
     */
    public Stream<ExtractedItem> extractShapes(Sheet sheet) {
        var drawing = sheet.getDrawingPatriarch();

        if (drawing == null) {
            return Stream.empty();
        }

        return switch (drawing) {
            case XSSFDrawing xssf -> extractXSSFShapes(xssf);
            case HSSFPatriarch hssf -> extractHSSFShapes(hssf);
            default -> Stream.empty();
        };
    }

    // --- Private helper methods ---

    /**
     * Internal record to hold cell address and value together.
     */
    private record CellValue(String address, String value) {
        boolean isEmpty() {
            return value.trim().isEmpty();
        }
    }

    /**
     * Extracts the value from a cell, handling different cell types.
     */
    private CellValue extractCellValue(Cell cell, DataFormatter dataFormatter) {
        var cellValue = switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> dataFormatter.formatCellValue(cell);
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            case FORMULA -> {
                try {
                    yield dataFormatter.formatCellValue(cell);
                } catch (Exception e) {
                    yield cell.getCellFormula();
                }
            }
            case BLANK -> "";
            default -> "";
        };

        return new CellValue(
            cell.getAddress().formatAsString(),
            cellValue.trim()
        );
    }

    /**
     * Extracts shapes from XSSF (xlsx/xlsm) drawings.
     */
    private Stream<ExtractedItem> extractXSSFShapes(XSSFDrawing drawing) {
        return drawing.getShapes().stream()
            .filter(XSSFSimpleShape.class::isInstance)
            .map(XSSFSimpleShape.class::cast)
            .map(XSSFSimpleShape::getText)
            .filter(text -> text != null && !text.trim().isEmpty())
            .map(text -> new ShapeItem(filePath, sheetName, text.trim()));
    }

    /**
     * Extracts shapes from HSSF (xls) drawings.
     */
    private Stream<ExtractedItem> extractHSSFShapes(HSSFPatriarch patriarch) {
        return patriarch.getChildren().stream()
            .filter(HSSFSimpleShape.class::isInstance)
            .map(HSSFSimpleShape.class::cast)
            .flatMap(this::extractHSSFShapeText)
            .filter(text -> !text.trim().isEmpty())
            .map(text -> new ShapeItem(filePath, sheetName, text.trim()));
    }

    /**
     * Safely extracts text from an HSSF shape, handling null cases.
     */
    private Stream<String> extractHSSFShapeText(HSSFSimpleShape shape) {
        try {
            return Optional.ofNullable(shape.getString())
                .map(HSSFRichTextString::getString)
                .stream();
        } catch (NullPointerException e) {
            // Skip shapes with null txtObjectRecord
            // This can happen with certain types of shapes in .xls files
            return Stream.empty();
        }
    }
}
