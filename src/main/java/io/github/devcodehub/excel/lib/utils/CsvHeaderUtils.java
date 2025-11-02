package io.github.devcodehub.excel.lib.utils;

import io.github.devcodehub.excel.lib.model.dto.example.csv.CsvHeaderBase;
import java.util.Arrays;

public class CsvHeaderUtils {
    public static <T extends Enum<T> & CsvHeaderBase> String[] getHeaders(Class<T> headerEnum) {
        return Arrays.stream(headerEnum.getEnumConstants())
                     .map(CsvHeaderBase::getDisplayName)
                     .toArray(String[]::new);
    }
}
