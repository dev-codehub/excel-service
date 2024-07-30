package com.excel.lib.model.dto.example;

import com.excel.lib.model.dto.excel.ExcelColor;
import com.excel.lib.model.dto.excel.ExcelHeaderBase;
import com.excel.lib.model.dto.excel.StyleDTO;
import lombok.AllArgsConstructor;

@AllArgsConstructor
public enum ExcelExampleHeader implements ExcelHeaderBase {
    VALUE1("value1", "Display Column 1", StyleDTO.builder().foregroundColor(ExcelColor.DARK_MINT).textColor(ExcelColor.WHITE).build()),
    VALUE2("value2", "Display Column 2", null),
    VALUE3("value3", "Display Column 3", null),
    VALUE4("value4", "Display Column 4", null);

    private final String field;
    private final String displayName;
    private final StyleDTO styles;

    @Override
    public String getField() {
        return field;
    }

    @Override
    public String getDisplayName() {
        return displayName;
    }

    @Override
    public StyleDTO getStyles() {
        return styles;
    }
}
