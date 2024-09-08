package io.github.devcodehub.excel.lib.utils;

import java.text.Format;
import java.text.SimpleDateFormat;
import java.util.Date;

public class DateUtils {
    public static final String DEFAULT_DATE_FORMAT = "yyyy/MM/dd";

    public static String getDateAsString(Date date) {
        return getDateAsString(date, DEFAULT_DATE_FORMAT);
    }

    public static String getDateAsString(Date date, String pattern) {
        if (date == null) {
            return null;
        }
        Format formatter = new SimpleDateFormat(pattern);
        return formatter.format(date);
    }
}
