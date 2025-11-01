package dev.untt.xlte;

import java.util.List;

/**
 * Formats extracted items in a human-readable format for terminal output.
 * Format: [SheetName:CellAddress] Content
 * Shows file headers when processing items from the same file.
 */
public class TerminalFormatter implements OutputFormatter {
    private String lastFilePath = null;

    @Override
    public String format(List<ExtractedItem> items) {
        if (items.isEmpty()) {
            return "";
        }

        var result = new StringBuilder();

        // Format each item
        for (var item : items) {
            var filePath = item.filePath();

            // Add file header if this is a different file from the last one
            if (!filePath.equals(lastFilePath)) {
                if (lastFilePath != null) {
                    // Add blank line between files (but not before the first file)
                    result.append("\n");
                }
                result.append("=== ").append(filePath).append(" ===\n");
                lastFilePath = filePath;
            }

            switch (item) {
                case CellItem cell -> {
                    result.append("[");
                    result.append(cell.sheetName());
                    result.append(":");
                    result.append(cell.cellAddress());
                    result.append("] ");
                    result.append(escapeText(cell.content()));
                    result.append("\n");
                }
                case ShapeItem shape -> {
                    result.append("[");
                    result.append(shape.sheetName());
                    result.append(":Shape] ");
                    result.append(escapeText(shape.content()));
                    result.append("\n");
                }
            }
        }

        return result.toString();
    }
}
