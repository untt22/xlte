package dev.untt.xlte;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class ExcelTextExtractor {

    public enum OutputMode {
        TERMINAL,  // Human-readable format for terminal
        TSV        // Tab-separated values for redirection/piping
    }

    private final OutputMode outputMode;

    public ExcelTextExtractor(OutputMode outputMode) {
        this.outputMode = outputMode;
    }

    /**
     * Escape newlines for text output (one cell per line format)
     */
    private String escapeText(String text) {
        if (text == null) {
            return "";
        }
        return text.replace("\r\n", "\\n")
                   .replace("\n", "\\n")
                   .replace("\r", "\\n");
    }

    public String extractText(File file) throws IOException {
        var result = new StringBuilder();
        var filePath = file.getPath();

        // Terminal mode: add file header
        if (outputMode == OutputMode.TERMINAL) {
            result.append("=== ").append(filePath).append(" ===\n");
        }

        try (var fis = new FileInputStream(file);
             var workbook = new XSSFWorkbook(fis)) {

            var numberOfSheets = workbook.getNumberOfSheets();

            for (var i = 0; i < numberOfSheets; i++) {
                var sheet = workbook.getSheetAt(i);
                var sheetName = sheet.getSheetName();

                // Extract text from cells
                result.append(extractCellText(filePath, sheetName, sheet));

                // Extract text from shapes
                result.append(extractShapeText(filePath, sheetName, sheet));
            }
        }

        // Terminal mode: add blank line after file
        if (outputMode == OutputMode.TERMINAL && !result.toString().trim().isEmpty()) {
            result.append("\n");
        }

        return result.toString().trim();
    }

    private String extractCellText(String filePath, String sheetName, XSSFSheet sheet) {
        var text = new StringBuilder();
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

                // Output non-empty cells only
                if (!cellValue.trim().isEmpty()) {
                    var cellAddress = cell.getAddress().formatAsString();

                    if (outputMode == OutputMode.TERMINAL) {
                        // Human-readable format: [Sheet:Cell] Content
                        text.append("[").append(sheetName).append(":")
                            .append(cellAddress).append("] ")
                            .append(cellValue).append("\n");
                    } else {
                        // TSV format: FilePath\tSheet\tCell\tContent
                        text.append(filePath).append("\t")
                            .append(sheetName).append("\t")
                            .append(cellAddress).append("\t")
                            .append(escapeText(cellValue)).append("\n");
                    }
                }
            }
        }

        return text.toString();
    }

    private String extractShapeText(String filePath, String sheetName, XSSFSheet sheet) {
        var text = new StringBuilder();

        var drawing = sheet.getDrawingPatriarch();
        if (drawing != null) {
            var shapes = drawing.getShapes();

            for (var shape : shapes) {
                if (shape instanceof XSSFSimpleShape simpleShape) {
                    var shapeText = simpleShape.getText();

                    if (shapeText != null && !shapeText.trim().isEmpty()) {
                        if (outputMode == OutputMode.TERMINAL) {
                            // Human-readable format: [Sheet:Shape] Content
                            text.append("[").append(sheetName).append(":Shape] ")
                                .append(shapeText.trim()).append("\n");
                        } else {
                            // TSV format: FilePath\tSheet\tShape\tContent
                            text.append(filePath).append("\t")
                                .append(sheetName).append("\t")
                                .append("Shape").append("\t")
                                .append(escapeText(shapeText.trim())).append("\n");
                        }
                    }
                }
            }
        }

        return text.toString();
    }
}
