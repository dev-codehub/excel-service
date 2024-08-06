package com.excel.lib.model.dto.excel.datatype;

import com.excel.lib.model.dto.excel.StyleDTO;
import com.excel.lib.utils.DateUtils;
import lombok.Builder;
import lombok.Getter;
import lombok.Setter;
import lombok.extern.jackson.Jacksonized;

import java.util.Date;

@Getter
@Setter
@Builder
@Jacksonized
public class DateExcel {
    private Date value;
    private String format;
    private StyleDTO styles;

    public static DateExcel fromValue(Date value) {
        return fromValue(value, DateUtils.DEFAULT_DATE_FORMAT, null);
    }

    public static DateExcel fromValue(Date value, String format) {
        return fromValue(value, format, null);
    }

    public static DateExcel fromValue(Date value, StyleDTO styles) {
        return fromValue(value, DateUtils.DEFAULT_DATE_FORMAT, styles);
    }

    public static DateExcel fromValue(Date value, String format, StyleDTO styles) {
        return DateExcel.builder()
                .value(value)
                .format(format)
                .styles(styles)
                .build();
    }
}
