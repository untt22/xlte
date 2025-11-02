package xlte;

/**
 * Represents a cell extracted from an Excel sheet.
 * Contains the file path, sheet name, cell address, and content.
 */
public record CellItem(
    String filePath,
    String sheetName,
    String cellAddress,
    String content
) implements ExtractedItem {
}
