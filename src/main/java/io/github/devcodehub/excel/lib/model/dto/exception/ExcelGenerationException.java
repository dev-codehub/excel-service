package io.github.devcodehub.excel.lib.model.dto.exception;

public class ExcelGenerationException extends Exception {
    private static final long serialVersionUID = 1L;

    public ExcelGenerationException(String message) {
        super(message);
    }

    public ExcelGenerationException(String message, Throwable cause) {
        super(message, cause);
    }
}
