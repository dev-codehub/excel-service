package com.excel.lib.model.dto.exception;

import lombok.AllArgsConstructor;
import lombok.Getter;

@Getter
@AllArgsConstructor
public enum ResponseCode {
    // Http code 200-299 - Successful responses
    SUCCESS(200, 0, "Success"),

    // Http code 400-499 - Client error responses
    BAD_REQUEST(400, 0, "Bad Request"),
    UNAUTHORIZED(401, 0, "Unauthorized"),
    FORBIDDEN(403, 0, "Forbidden"),
    NOT_FOUND(404, 0, "Not Found"),
    REQUEST_TIMEOUT(408, 0, "Request Timeout"),

    // Http code 500-599 - Server error responses
    INTERNAL_SERVER_ERROR(500, 0, "Internal Server Error"),
    ERROR_GENERATING_DYNAMIC_EXCEL(500, 15, "Error generating dynamic excel"),
    NOT_IMPLEMENTED(501, 0, "Not Implemented"),
    UNAVAILABLE(503, 0, "Unavailable");

    private final int httpCode;
    private final int internalCode;
    private final String message;
}
