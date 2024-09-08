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
public class ExcelSettings {
    private String sheetName;
    @Builder.Default
    private ExcelCustomStyles excelCustomStyles = ExcelCustomStyles.builder().build();
    @Builder.Default
    private boolean headerFilterActive = false;
    @Builder.Default
    private int headerFilterOffset = 0;
    private FreezePane freezePane;
    @Builder.Default
    private int rowOffset = 0;
    @Builder.Default
    private int colOffset = 0;

    @Getter
    @Setter
    @Builder
    @NoArgsConstructor
    @AllArgsConstructor
    public static class FreezePane {
        @Builder.Default
        private int colSplit = 0;
        @Builder.Default
        private int rowSplit = 0;
        private Integer leftmostColumn;
        private Integer topRow;

        public int getLeftmostColumn() {
            return leftmostColumn != null ? leftmostColumn : colSplit;
        }

        public int getTopRow() {
            return topRow != null ? topRow : rowSplit;
        }
    }
}
