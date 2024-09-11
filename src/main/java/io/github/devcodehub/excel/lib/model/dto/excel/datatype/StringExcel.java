package io.github.devcodehub.excel.lib.model.dto.excel.datatype;

import io.github.devcodehub.excel.lib.model.dto.excel.StyleDTO;
import lombok.Builder;
import lombok.Getter;
import lombok.Setter;
import lombok.extern.jackson.Jacksonized;

@Getter
@Setter
@Builder
@Jacksonized
public class StringExcel {
    private String value;
    private StyleDTO styles;

    public static StringExcel fromValue(String value) {
        return fromValue(value, null);
    }

    public static StringExcel fromValue(String value, StyleDTO styles) {
        return StringExcel.builder()
                .value(value)
                .styles(styles)
                .build();
    }
}
