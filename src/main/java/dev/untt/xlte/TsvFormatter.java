package dev.untt.xlte;

import java.util.List;

/**
 * Formats extracted items as Tab-Separated Values (TSV) for piping/redirection.
 * Format: FilePath\tSheetName\tCellAddress\tContent
 */
public class TsvFormatter implements OutputFormatter {

    @Override
    public String format(List<ExtractedItem> items) {
        if (items.isEmpty()) {
            return "";
        }

        var result = new StringBuilder();

        // Format each item as TSV (always includes file path)
        for (var item : items) {
            switch (item) {
                case CellItem cell -> {
                    result.append(cell.filePath());
                    result.append("\t");
                    result.append(cell.sheetName());
                    result.append("\t");
                    result.append(cell.cellAddress());
                    result.append("\t");
                    result.append(escapeText(cell.content()));
                    result.append("\n");
                }
                case ShapeItem shape -> {
                    result.append(shape.filePath());
                    result.append("\t");
                    result.append(shape.sheetName());
                    result.append("\t");
                    result.append("Shape");
                    result.append("\t");
                    result.append(escapeText(shape.content()));
                    result.append("\n");
                }
            }
        }

        return result.toString();
    }
}
