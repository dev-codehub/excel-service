package com.excel.lib.model.dto.excel.datatype;

import com.excel.lib.model.dto.excel.StyleDTO;
import lombok.Builder;
import lombok.Getter;
import lombok.Setter;
import lombok.extern.jackson.Jacksonized;

@Getter
@Setter
@Builder
@Jacksonized
public class Number {
    private Double value;
    private String format;
    private StyleDTO styles;

    public static Number fromValue(Double value) {
        return fromValue(value, null, null);
    }

    public static Number fromValue(Double value, String format) {
        return fromValue(value, format, null);
    }

    public static Number fromValue(Double value, StyleDTO styles) {
        return fromValue(value, null, styles);
    }

    public static Number fromValue(Double value, String format, StyleDTO styles) {
        return Number.builder()
                .value(value)
                .format(format)
                .styles(styles)
                .build();
    }
}
