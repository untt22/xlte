package xlte;

import java.util.List;
import java.util.stream.Collectors;

/**
 * Formats extracted items as Tab-Separated Values (TSV) for piping/redirection.
 * Format: FilePath\tSheetName\tCellAddress\tContent
 *
 * This formatter is stateless and uses functional stream operations.
 */
public class TsvFormatter implements OutputFormatter {

    @Override
    public String format(List<ExtractedItem> items) {
        if (items.isEmpty()) {
            return "";
        }

        return items.stream()
            .map(this::formatItemAsTsv)
            .collect(Collectors.joining());
    }

    /**
     * Formats a single item as a TSV line.
     * Pure function with no side effects.
     */
    private String formatItemAsTsv(ExtractedItem item) {
        return switch (item) {
            case CellItem cell -> String.join("\t",
                cell.filePath(),
                cell.sheetName(),
                cell.cellAddress(),
                escapeText(cell.content())
            ) + "\n";

            case ShapeItem shape -> String.join("\t",
                shape.filePath(),
                shape.sheetName(),
                "Shape",
                escapeText(shape.content())
            ) + "\n";
        };
    }
}
