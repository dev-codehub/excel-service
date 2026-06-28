package io.github.devcodehub.excel.lib.service;

import io.github.devcodehub.excel.lib.model.dto.excel.ExcelColor;
import io.github.devcodehub.excel.lib.model.dto.excel.ExcelCustomStyles;
import io.github.devcodehub.excel.lib.model.dto.excel.StyleDTO;
import io.github.devcodehub.excel.lib.model.dto.excel.datatype.DateExcel;
import io.github.devcodehub.excel.lib.model.dto.excel.datatype.Merge;
import io.github.devcodehub.excel.lib.model.dto.excel.datatype.Number;
import io.github.devcodehub.excel.lib.model.dto.excel.datatype.StringExcel;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertSame;

class StyleServiceTest {

    private Workbook workbook;
    private ExcelCustomStyles defaultCustomStyles;

    @BeforeEach
    void setUp() {
        workbook = new XSSFWorkbook();
        defaultCustomStyles = ExcelCustomStyles.builder().build();
    }

    @AfterEach
    void tearDown() throws Exception {
        workbook.close();
    }

    @Test
    void constructor_headerCellStyle_isNotNull() {
        StyleService service = new StyleService(workbook, defaultCustomStyles);

        assertNotNull(service.getHeaderCellStyle());
    }

    @Test
    void constructor_dataCellStyle_isNotNull() {
        StyleService service = new StyleService(workbook, defaultCustomStyles);

        assertNotNull(service.getDataCellStyle());
    }

    @Test
    void getHeaderCellStyle_withNullStyleDTO_returnsSameDefaultInstance() {
        StyleService service = new StyleService(workbook, defaultCustomStyles);

        CellStyle result = service.getHeaderCellStyle(workbook, null);

        assertSame(service.getHeaderCellStyle(), result);
    }

    @Test
    void getCellStyle_forStringExcel_withNullStyleDTO_returnsNonNull() {
        StyleService service = new StyleService(workbook, defaultCustomStyles);

        assertNotNull(service.getCellStyle(workbook, StringExcel.class, null));
    }

    @Test
    void getCellStyle_forNumber_withNullStyleDTO_returnsNonNull() {
        StyleService service = new StyleService(workbook, defaultCustomStyles);

        assertNotNull(service.getCellStyle(workbook, Number.class, null));
    }

    @Test
    void getCellStyle_forDateExcel_withNullStyleDTO_returnsNonNull() {
        StyleService service = new StyleService(workbook, defaultCustomStyles);

        assertNotNull(service.getCellStyle(workbook, DateExcel.class, null));
    }

    @Test
    void getCellStyle_forMerge_withNullStyleDTO_returnsNonNull() {
        StyleService service = new StyleService(workbook, defaultCustomStyles);

        assertNotNull(service.getCellStyle(workbook, Merge.class, null));
    }

    @Test
    void getNewCellStyle_bold_appliedToFont() {
        StyleService service = new StyleService(workbook, defaultCustomStyles);
        StyleDTO styleDTO = StyleDTO.builder().bold(true).build();

        CellStyle style = service.getCellStyle(workbook, StringExcel.class, styleDTO);

        Font font = workbook.getFontAt(style.getFontIndex());
        assertEquals(true, font.getBold());
    }

    @Test
    void getNewCellStyle_fontHeight_appliedToFont() {
        StyleService service = new StyleService(workbook, defaultCustomStyles);
        StyleDTO styleDTO = StyleDTO.builder().fontHeight(16.0).build();

        CellStyle style = service.getCellStyle(workbook, StringExcel.class, styleDTO);

        Font font = workbook.getFontAt(style.getFontIndex());
        assertEquals(16, font.getFontHeightInPoints());
    }

    @Test
    void getNewCellStyle_horizontalAlignment_applied() {
        StyleService service = new StyleService(workbook, defaultCustomStyles);
        StyleDTO styleDTO = StyleDTO.builder().horizontalAlignment(HorizontalAlignment.LEFT).build();

        CellStyle style = service.getCellStyle(workbook, StringExcel.class, styleDTO);

        assertEquals(HorizontalAlignment.LEFT, style.getAlignment());
    }

    @Test
    void getNewCellStyle_verticalAlignment_applied() {
        StyleService service = new StyleService(workbook, defaultCustomStyles);
        StyleDTO styleDTO = StyleDTO.builder().verticalAlignment(VerticalAlignment.TOP).build();

        CellStyle style = service.getCellStyle(workbook, StringExcel.class, styleDTO);

        assertEquals(VerticalAlignment.TOP, style.getVerticalAlignment());
    }

    @Test
    void getNewCellStyle_foregroundColor_setsSolidForeground() {
        StyleService service = new StyleService(workbook, defaultCustomStyles);
        StyleDTO styleDTO = StyleDTO.builder().foregroundColor(ExcelColor.LIGHT_BLUE).build();

        CellStyle style = service.getCellStyle(workbook, StringExcel.class, styleDTO);

        assertEquals(FillPatternType.SOLID_FOREGROUND, style.getFillPattern());
    }

    @Test
    void defaultCustomStyles_headerBorderActive_producesNonNoneBorder() {
        ExcelCustomStyles styles = ExcelCustomStyles.builder().headerBorderActive(true).build();
        StyleService service = new StyleService(workbook, styles);

        assertNotEquals(BorderStyle.NONE, service.getHeaderCellStyle().getBorderTop());
    }

    @Test
    void defaultCustomStyles_customFontFamily_reflectedInFont() {
        ExcelCustomStyles styles = ExcelCustomStyles.builder().fontFamily("Times New Roman").build();
        StyleService service = new StyleService(workbook, styles);

        Font font = workbook.getFontAt(service.getHeaderCellStyle().getFontIndex());
        assertEquals("Times New Roman", font.getFontName());
    }
}
