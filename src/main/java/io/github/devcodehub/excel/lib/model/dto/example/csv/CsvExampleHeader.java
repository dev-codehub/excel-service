package io.github.devcodehub.excel.lib.model.dto.example.csv;

import io.github.devcodehub.excel.lib.model.dto.csv.CsvHeaderBase;
import lombok.AllArgsConstructor;

@AllArgsConstructor
public enum CsvExampleHeader implements CsvHeaderBase {
    VALUE1("value1", "Display Column 1"),
    VALUE2("value2", "Display Column 2"),
    VALUE3("value3", "Display Column 3"),
    VALUE4("value4", "Display Column 4"),
    ;

    private final String field;
    private final String displayName;

    @Override
    public String getField() {
        return field;
    }

    @Override
    public String getDisplayName() {
        return displayName;
    }
}
