package xlte;

/**
 * Represents a shape (text box, etc.) extracted from an Excel sheet.
 * Contains the file path, sheet name, and content.
 */
public record ShapeItem(
    String filePath,
    String sheetName,
    String content
) implements ExtractedItem {
}
