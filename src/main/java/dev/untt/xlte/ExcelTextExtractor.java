package dev.untt.xlte;

import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFSimpleShape;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * Extracts text content from Excel files (.xls, .xlsx, .xlsm).
 * This class is responsible for extraction only, not formatting.
 */
public class ExcelTextExtractor {

    /**
     * Extracts all text content from an Excel file.
     *
     * @param file The Excel file to extract from
     * @return A list of extracted items (cells and shapes)
     * @throws IOException If the file cannot be read
     */
    public List<ExtractedItem> extract(File file) throws IOException {
        var items = new ArrayList<ExtractedItem>();
        var filePath = file.getPath();

        try (var fis = new FileInputStream(file);
             var workbook = WorkbookFactory.create(fis)) {

            var numberOfSheets = workbook.getNumberOfSheets();

            for (var i = 0; i < numberOfSheets; i++) {
                var sheet = workbook.getSheetAt(i);
                var sheetName = sheet.getSheetName();

                // Extract text from cells
                items.addAll(extractCells(filePath, sheetName, sheet));

                // Extract text from shapes
                items.addAll(extractShapes(filePath, sheetName, sheet));
            }
        }

        return items;
    }

    /**
     * Extracts all non-empty cells from a sheet.
     */
    private List<ExtractedItem> extractCells(String filePath, String sheetName, Sheet sheet) {
        var items = new ArrayList<ExtractedItem>();
        var dataFormatter = new DataFormatter();

        for (var row : sheet) {
            for (var cell : row) {
                if (cell == null) {
                    continue;
                }

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

                // Only add non-empty cells
                if (!cellValue.trim().isEmpty()) {
                    var cellAddress = cell.getAddress().formatAsString();
                    items.add(new CellItem(filePath, sheetName, cellAddress, cellValue));
                }
            }
        }

        return items;
    }

    /**
     * Extracts all text from shapes (text boxes, etc.) in a sheet.
     */
    private List<ExtractedItem> extractShapes(String filePath, String sheetName, Sheet sheet) {
        var items = new ArrayList<ExtractedItem>();

        var drawing = sheet.getDrawingPatriarch();
        if (drawing != null) {
            // Handle XSSF (xlsx/xlsm) shapes
            if (drawing instanceof XSSFDrawing xssfDrawing) {
                var shapes = xssfDrawing.getShapes();
                for (var shape : shapes) {
                    if (shape instanceof XSSFSimpleShape simpleShape) {
                        var shapeText = simpleShape.getText();
                        if (shapeText != null && !shapeText.trim().isEmpty()) {
                            items.add(new ShapeItem(filePath, sheetName, shapeText.trim()));
                        }
                    }
                }
            }
            // Handle HSSF (xls) shapes
            else if (drawing instanceof HSSFPatriarch hssfPatriarch) {
                var shapes = hssfPatriarch.getChildren();
                for (var shape : shapes) {
                    if (shape instanceof HSSFSimpleShape simpleShape) {
                        try {
                            var shapeText = simpleShape.getString();
                            if (shapeText != null && !shapeText.getString().trim().isEmpty()) {
                                items.add(new ShapeItem(filePath, sheetName, shapeText.getString().trim()));
                            }
                        } catch (NullPointerException e) {
                            // Skip shapes with null txtObjectRecord
                            // This can happen with certain types of shapes in .xls files
                        }
                    }
                }
            }
        }

        return items;
    }
}
