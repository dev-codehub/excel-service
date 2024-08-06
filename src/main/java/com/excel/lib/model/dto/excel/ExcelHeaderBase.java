package com.excel.lib.model.dto.excel;

public interface ExcelHeaderBase {
    String getField();

    String getDisplayName();

    default StyleDTO getStyles() {
        return new StyleDTO();
    }
}
