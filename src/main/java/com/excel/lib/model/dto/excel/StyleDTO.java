package com.excel.lib.model.dto.excel;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

/**
 * This class represents all styles for a cell in Excel
 */
@Getter
@Setter
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class StyleDTO {
    private Boolean bold;
    private Double fontHeight;
    private ExcelColorBase foregroundColor;
    private ExcelColorBase textColor;
    private VerticalAlignment verticalAlignment;
    private HorizontalAlignment horizontalAlignment;
    private Integer minWidth;
    private Integer maxWidth;
}
