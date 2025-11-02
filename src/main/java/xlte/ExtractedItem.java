package xlte;

/**
 * Represents an item extracted from an Excel file.
 * This is a sealed interface that can only be implemented by CellItem and ShapeItem.
 */
public sealed interface ExtractedItem permits CellItem, ShapeItem {
    /**
     * @return The path to the Excel file this item was extracted from
     */
    String filePath();

    /**
     * @return The name of the sheet this item was extracted from
     */
    String sheetName();

    /**
     * @return The text content of this item
     */
    String content();
}
