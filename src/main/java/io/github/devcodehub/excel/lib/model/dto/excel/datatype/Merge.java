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
public class Merge {
    private int range;
    private int offset;
    private String value;
    private Orientation orientation;
    private StyleDTO styles;

    public static Merge fromValue(int range, int offset, String value, Orientation orientation) {
        return fromValue(range, offset, value, orientation, null);
    }

    public static Merge fromValue(int range, int offset, String value, Orientation orientation, StyleDTO style) {
        return Merge.builder()
                .range(range)
                .offset(offset)
                .value(value)
                .orientation(orientation)
                .styles(style)
                .build();
    }

    public enum Orientation {
        VERTICAL,
        HORIZONTAL
    }
}
