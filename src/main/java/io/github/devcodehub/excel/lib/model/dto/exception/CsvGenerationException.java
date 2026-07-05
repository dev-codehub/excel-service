package io.github.devcodehub.excel.lib.model.dto.exception;

public class CsvGenerationException extends Exception {
    private static final long serialVersionUID = 1L;

    public CsvGenerationException(String message) {
        super(message);
    }

    public CsvGenerationException(String message, Throwable cause) {
        super(message, cause);
    }
}
