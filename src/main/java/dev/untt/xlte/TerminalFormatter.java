package dev.untt.xlte;

import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

/**
 * Formats extracted items in a human-readable format for terminal output.
 * Format: [SheetName:CellAddress] Content
 * Shows file headers when processing items from the same file.
 *
 * This formatter is stateless and thread-safe.
 */
public class TerminalFormatter implements OutputFormatter {

    @Override
    public String format(List<ExtractedItem> items) {
        if (items.isEmpty()) {
            return "";
        }

        // Group items by file path (preserving order)
        Map<String, List<ExtractedItem>> groupedByFile = items.stream()
            .collect(Collectors.groupingBy(
                ExtractedItem::filePath,
                LinkedHashMap::new,  // Preserve insertion order
                Collectors.toList()
            ));

        // Format each file group
        return groupedByFile.entrySet().stream()
            .map(entry -> formatFileGroup(entry.getKey(), entry.getValue()))
            .collect(Collectors.joining("\n"));
    }

    /**
     * Formats a group of items from the same file.
     * Pure function with no side effects.
     */
    private String formatFileGroup(String filePath, List<ExtractedItem> items) {
        var header = "=== " + filePath + " ===\n";
        var itemsFormatted = items.stream()
            .map(this::formatItem)
            .collect(Collectors.joining());

        return header + itemsFormatted;
    }

    /**
     * Formats a single item.
     * Pure function with no side effects.
     */
    private String formatItem(ExtractedItem item) {
        return switch (item) {
            case CellItem cell -> String.format(
                "[%s:%s] %s\n",
                cell.sheetName(),
                cell.cellAddress(),
                escapeText(cell.content())
            );
            case ShapeItem shape -> String.format(
                "[%s:Shape] %s\n",
                shape.sheetName(),
                escapeText(shape.content())
            );
        };
    }
}
