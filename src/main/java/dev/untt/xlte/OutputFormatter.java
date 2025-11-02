package dev.untt.xlte;

import java.util.List;

/**
 * Interface for formatting extracted Excel items for output.
 * Implementations can format the data in different ways (terminal, TSV, etc.).
 */
public interface OutputFormatter {
    /**
     * Formats a list of extracted items into a string for output.
     *
     * @param items The list of extracted items to format
     * @return The formatted string ready for output
     */
    String format(List<ExtractedItem> items);

    /**
     * Escapes special characters in text for safe output.
     * Converts newlines (\r\n, \n, \r) to spaces.
     *
     * @param text The text to escape
     * @return The escaped text, or empty string if text is null
     */
    default String escapeText(String text) {
        if (text == null) {
            return "";
        }
        return text.replace("\r\n", " ")
                   .replace("\n", " ")
                   .replace("\r", " ");
    }
}
