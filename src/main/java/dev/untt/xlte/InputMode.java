package dev.untt.xlte;

import java.io.File;

/**
 * Represents the input mode for the application.
 * Either processing a single file or a directory.
 */
sealed interface InputMode permits InputMode.SingleFile, InputMode.Directory {

    /**
     * Single file mode: process one Excel file.
     */
    record SingleFile(File file) implements InputMode {}

    /**
     * Directory mode: process all Excel files in a directory.
     */
    record Directory(File dir, boolean recursive) implements InputMode {}
}
