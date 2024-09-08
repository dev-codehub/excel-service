package io.github.devcodehub.excel.lib.model.dto.excel;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@Getter
@Setter
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class ExcelCustomStyles {
    @Builder.Default
    private short headerFontHeight = 11;
    @Builder.Default
    private Double headersHeight = 24.0;
    @Builder.Default
    private boolean headerBold = false;
    @Builder.Default
    private boolean headerBorderActive = true;
    @Builder.Default
    private ExcelColorBase headerBorderColor = ExcelColor.BLACK;
    @Builder.Default
    private short dataCellFontHeight = 11;
    @Builder.Default
    private boolean dataCellBold = false;
    @Builder.Default
    private boolean dataBorderActive = false;
    @Builder.Default
    private ExcelColorBase dataBorderColor = ExcelColor.LIGHT_GREY;
    @Builder.Default
    public String fontFamily = "Calibri";

    public void setHeaderBorderColor(ExcelColorBase headerBorderColor) {
        this.headerBorderColor = headerBorderColor;
        this.headerBorderActive = true;
    }

    public void setDataBorderColor(ExcelColorBase dataBorderColor) {
        this.dataBorderColor = dataBorderColor;
        this.dataBorderActive = true;
    }
}
