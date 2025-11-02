package io.github.devcodehub.excel.lib.model.dto.example.csv;

import lombok.AllArgsConstructor;

@AllArgsConstructor
public enum CsvExampleHeader implements CsvHeaderBase {
    VALUE1("Display Column 1"),
    VALUE2("Display Column 2"),
    VALUE3("Display Column 3"),
    VALUE4("Display Column 4"),
    ;

    private final String displayName;

    @Override
    public String getDisplayName() {
        return displayName;
    }
}
