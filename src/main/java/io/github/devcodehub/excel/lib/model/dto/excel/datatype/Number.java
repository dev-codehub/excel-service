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
public class Number {
    private Double value;
    private StyleDTO styles;

    public static Number fromValue(Double value) {
        return fromValue(value, null);
    }

    public static Number fromValue(Double value, StyleDTO styles) {
        return Number.builder()
                .value(value)
                .styles(styles)
                .build();
    }
}
