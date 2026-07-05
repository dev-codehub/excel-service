package io.github.devcodehub.excel.lib.model.dto.excel;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import java.util.Collections;
import java.util.List;

@Getter
@Setter
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class ExcelReadSettings {
    private String sheetName;

    @Builder.Default
    private int rowOffset = 0;

    @Builder.Default
    private int colOffset = 0;

    /**
     * Header display names used to locate the table. When non-empty, the reader scans the sheet
     * row by row until it finds a row containing all specified keys, and uses that as the header row.
     * When empty, rowOffset is used as the header row index directly.
     */
    @Builder.Default
    private List<String> searchKeys = Collections.emptyList();

    /**
     * Maximum number of rows to scan when using searchKeys (0 = no limit, scan entire sheet).
     */
    @Builder.Default
    private int maxScanRows = 0;
}
