package io.github.devcodehub.excel.lib.utils;

import io.github.devcodehub.excel.lib.model.dto.csv.CsvHeaderBase;

import java.util.List;

public class CsvHeaderUtils {
    private CsvHeaderUtils() {}

    public static String[] getDisplayNames(List<? extends CsvHeaderBase> headers) {
        return headers.stream()
                      .map(CsvHeaderBase::getDisplayName)
                      .toArray(String[]::new);
    }
}
