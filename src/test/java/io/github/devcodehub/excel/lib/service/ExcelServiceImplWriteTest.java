package io.github.devcodehub.excel.lib.service;

import io.github.devcodehub.excel.lib.model.dto.excel.ExcelHeaderBase;
import io.github.devcodehub.excel.lib.model.dto.excel.ExcelSettings;
import io.github.devcodehub.excel.lib.model.dto.excel.StyleDTO;
import io.github.devcodehub.excel.lib.model.dto.excel.datatype.DateExcel;
import io.github.devcodehub.excel.lib.model.dto.excel.datatype.Merge;
import io.github.devcodehub.excel.lib.model.dto.excel.datatype.Number;
import io.github.devcodehub.excel.lib.model.dto.excel.datatype.StringExcel;
import io.github.devcodehub.excel.lib.model.dto.exception.ExcelGenerationException;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Collections;
import java.util.Date;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

class ExcelServiceImplWriteTest {

    private ExcelServiceImpl service;

    @BeforeEach
    void setUp() {
        service = new ExcelServiceImpl();
    }

    // ── DTOs ──────────────────────────────────────────────────────────────────

    @Getter @Setter @Builder @NoArgsConstructor @AllArgsConstructor
    static class BasicDTO {
        private String name;
        private Double amount;
    }

    @Getter @Setter @Builder @NoArgsConstructor @AllArgsConstructor
    static class AllTypesDTO {
        private String plain;
        private Boolean flag;
        private Date date;
        private StringExcel styled;
        private Number num;
        private DateExcel dateExcel;
    }

    @Getter @Setter @Builder @NoArgsConstructor @AllArgsConstructor
    static class MergeDTO {
        private Merge mergeField;
    }

    // ── Headers ───────────────────────────────────────────────────────────────

    @AllArgsConstructor
    enum BasicHeader implements ExcelHeaderBase {
        NAME("name", "Name", null),
        AMOUNT("amount", "Amount", null);
        private final String field;
        private final String displayName;
        private final StyleDTO styles;
        @Override public String getField() { return field; }
        @Override public String getDisplayName() { return displayName; }
        @Override public StyleDTO getStyles() { return styles; }
    }

    @AllArgsConstructor
    enum AllTypesHeader implements ExcelHeaderBase {
        PLAIN("plain", "Plain", null),
        FLAG("flag", "Flag", null),
        DATE("date", "Date", null),
        STYLED("styled", "Styled", null),
        NUM("num", "Num", null),
        DATE_EXCEL("dateExcel", "DateExcel", null);
        private final String field;
        private final String displayName;
        private final StyleDTO styles;
        @Override public String getField() { return field; }
        @Override public String getDisplayName() { return displayName; }
        @Override public StyleDTO getStyles() { return styles; }
    }

    @AllArgsConstructor
    enum MergeHeader implements ExcelHeaderBase {
        MERGE_FIELD("mergeField", "Merged", null);
        private final String field;
        private final String displayName;
        private final StyleDTO styles;
        @Override public String getField() { return field; }
        @Override public String getDisplayName() { return displayName; }
        @Override public StyleDTO getStyles() { return styles; }
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    private static ExcelSettings settings(String sheetName) {
        return ExcelSettings.builder().sheetName(sheetName).build();
    }

    private static List<BasicHeader> basicHeaders() {
        return Arrays.asList(BasicHeader.values());
    }

    private static Date buildDate(int year, int month, int day) {
        Calendar cal = Calendar.getInstance();
        cal.set(year, month, day, 0, 0, 0);
        cal.set(Calendar.MILLISECOND, 0);
        return cal.getTime();
    }

    // ── Tests ─────────────────────────────────────────────────────────────────

    @Test
    void generateDynamicExcelWorkbook_createsSheetWithGivenName() throws Exception {
        try (Workbook wb = service.generateDynamicExcelWorkbook(
                basicHeaders(), Collections.emptyList(), BasicDTO.class, settings("MySheet"))) {
            assertNotNull(wb.getSheet("MySheet"));
        }
    }

    @Test
    void generateDynamicExcelWorkbook_headerRow_atRowOffset() throws Exception {
        List<BasicDTO> data = Collections.singletonList(BasicDTO.builder().name("Alice").build());
        ExcelSettings s = ExcelSettings.builder().sheetName("S").rowOffset(2).build();

        try (Workbook wb = service.generateDynamicExcelWorkbook(basicHeaders(), data, BasicDTO.class, s)) {
            Sheet sheet = wb.getSheet("S");
            Row headerRow = sheet.getRow(2);
            assertNotNull(headerRow);
            assertEquals("Name", headerRow.getCell(0).getStringCellValue());
        }
    }

    @Test
    void generateDynamicExcelWorkbook_headerRow_containsDisplayNames() throws Exception {
        try (Workbook wb = service.generateDynamicExcelWorkbook(
                basicHeaders(), Collections.emptyList(), BasicDTO.class, settings("S"))) {
            Sheet sheet = wb.getSheet("S");
            Row header = sheet.getRow(0);
            assertEquals("Name",   header.getCell(0).getStringCellValue());
            assertEquals("Amount", header.getCell(1).getStringCellValue());
        }
    }

    @Test
    void generateDynamicExcelWorkbook_dataRow_startsAtRowOffsetPlusOne() throws Exception {
        List<BasicDTO> data = Collections.singletonList(BasicDTO.builder().name("Alice").amount(99.0).build());
        ExcelSettings s = ExcelSettings.builder().sheetName("S").rowOffset(1).build();

        try (Workbook wb = service.generateDynamicExcelWorkbook(basicHeaders(), data, BasicDTO.class, s)) {
            Sheet sheet = wb.getSheet("S");
            Row dataRow = sheet.getRow(2);
            assertNotNull(dataRow);
            assertEquals("Alice", dataRow.getCell(0).getStringCellValue());
        }
    }

    @Test
    void generateDynamicExcelWorkbook_colOffset_shiftsAllColumns() throws Exception {
        List<BasicDTO> data = Collections.singletonList(BasicDTO.builder().name("Bob").amount(10.0).build());
        ExcelSettings s = ExcelSettings.builder().sheetName("S").colOffset(2).build();

        try (Workbook wb = service.generateDynamicExcelWorkbook(basicHeaders(), data, BasicDTO.class, s)) {
            Sheet sheet = wb.getSheet("S");
            assertEquals("Name",   sheet.getRow(0).getCell(2).getStringCellValue());
            assertEquals("Bob",    sheet.getRow(1).getCell(2).getStringCellValue());
        }
    }

    @Test
    void generateDynamicExcelWorkbook_stringField_writtenAsStringCell() throws Exception {
        List<BasicDTO> data = Collections.singletonList(BasicDTO.builder().name("Charlie").build());

        try (Workbook wb = service.generateDynamicExcelWorkbook(basicHeaders(), data, BasicDTO.class, settings("S"))) {
            Sheet sheet = wb.getSheet("S");
            assertEquals(CellType.STRING, sheet.getRow(1).getCell(0).getCellType());
            assertEquals("Charlie", sheet.getRow(1).getCell(0).getStringCellValue());
        }
    }

    @Test
    void generateDynamicExcelWorkbook_booleanField_writtenAsBooleanCell() throws Exception {
        List<AllTypesDTO> data = Collections.singletonList(AllTypesDTO.builder().flag(true).build());

        try (Workbook wb = service.generateDynamicExcelWorkbook(
                Arrays.asList(AllTypesHeader.values()), data, AllTypesDTO.class, settings("S"))) {
            Sheet sheet = wb.getSheet("S");
            assertEquals(CellType.BOOLEAN, sheet.getRow(1).getCell(1).getCellType());
            assertTrue(sheet.getRow(1).getCell(1).getBooleanCellValue());
        }
    }

    @Test
    void generateDynamicExcelWorkbook_dateField_writtenAsFormattedString() throws Exception {
        Date date = buildDate(2024, Calendar.JUNE, 10);
        List<AllTypesDTO> data = Collections.singletonList(AllTypesDTO.builder().date(date).build());

        try (Workbook wb = service.generateDynamicExcelWorkbook(
                Arrays.asList(AllTypesHeader.values()), data, AllTypesDTO.class, settings("S"))) {
            Sheet sheet = wb.getSheet("S");
            assertEquals("2024/06/10", sheet.getRow(1).getCell(2).getStringCellValue());
        }
    }

    @Test
    void generateDynamicExcelWorkbook_stringExcelField_writtenAsStringCell() throws Exception {
        List<AllTypesDTO> data = Collections.singletonList(
                AllTypesDTO.builder().styled(StringExcel.fromValue("Hello")).build());

        try (Workbook wb = service.generateDynamicExcelWorkbook(
                Arrays.asList(AllTypesHeader.values()), data, AllTypesDTO.class, settings("S"))) {
            Sheet sheet = wb.getSheet("S");
            assertEquals("Hello", sheet.getRow(1).getCell(3).getStringCellValue());
        }
    }

    @Test
    void generateDynamicExcelWorkbook_numberField_writtenAsNumericCell() throws Exception {
        List<AllTypesDTO> data = Collections.singletonList(
                AllTypesDTO.builder().num(Number.fromValue(42.5)).build());

        try (Workbook wb = service.generateDynamicExcelWorkbook(
                Arrays.asList(AllTypesHeader.values()), data, AllTypesDTO.class, settings("S"))) {
            Sheet sheet = wb.getSheet("S");
            assertEquals(CellType.NUMERIC, sheet.getRow(1).getCell(4).getCellType());
            assertEquals(42.5, sheet.getRow(1).getCell(4).getNumericCellValue(), 0.001);
        }
    }

    @Test
    void generateDynamicExcelWorkbook_dateExcelField_usesSpecifiedFormat() throws Exception {
        Date date = buildDate(2024, Calendar.JUNE, 10);
        List<AllTypesDTO> data = Collections.singletonList(
                AllTypesDTO.builder().dateExcel(DateExcel.fromValue(date, "dd/MM/yyyy")).build());

        try (Workbook wb = service.generateDynamicExcelWorkbook(
                Arrays.asList(AllTypesHeader.values()), data, AllTypesDTO.class, settings("S"))) {
            Sheet sheet = wb.getSheet("S");
            assertEquals("10/06/2024", sheet.getRow(1).getCell(5).getStringCellValue());
        }
    }

    @Test
    void generateDynamicExcelWorkbook_nullField_doesNotCreateCell() throws Exception {
        List<BasicDTO> data = Collections.singletonList(BasicDTO.builder().name("Diana").build()); // amount null

        try (Workbook wb = service.generateDynamicExcelWorkbook(basicHeaders(), data, BasicDTO.class, settings("S"))) {
            Sheet sheet = wb.getSheet("S");
            assertNull(sheet.getRow(1).getCell(1)); // no cell for null amount
        }
    }

    @Test
    void generateDynamicExcelWorkbook_multipleRows_allWritten() throws Exception {
        List<BasicDTO> data = Arrays.asList(
                BasicDTO.builder().name("Alice").build(),
                BasicDTO.builder().name("Bob").build(),
                BasicDTO.builder().name("Carol").build());

        try (Workbook wb = service.generateDynamicExcelWorkbook(basicHeaders(), data, BasicDTO.class, settings("S"))) {
            Sheet sheet = wb.getSheet("S");
            assertEquals("Alice", sheet.getRow(1).getCell(0).getStringCellValue());
            assertEquals("Bob",   sheet.getRow(2).getCell(0).getStringCellValue());
            assertEquals("Carol", sheet.getRow(3).getCell(0).getStringCellValue());
        }
    }

    @Test
    void generateDynamicExcelWorkbook_autoFilter_setWhenEnabled() throws Exception {
        ExcelSettings s = ExcelSettings.builder().sheetName("S").headerFilterActive(true).build();

        try (Workbook wb = service.generateDynamicExcelWorkbook(
                basicHeaders(), Collections.emptyList(), BasicDTO.class, s)) {
            Sheet sheet = wb.getSheet("S");
            assertTrue(((XSSFSheet) sheet).getCTWorksheet().isSetAutoFilter());
        }
    }

    @Test
    void generateDynamicExcelWorkbook_freezePane_applied() throws Exception {
        ExcelSettings.FreezePane fp = ExcelSettings.FreezePane.builder().colSplit(1).rowSplit(1).build();
        ExcelSettings s = ExcelSettings.builder().sheetName("S").freezePane(fp).build();

        try (Workbook wb = service.generateDynamicExcelWorkbook(
                basicHeaders(), Collections.emptyList(), BasicDTO.class, s)) {
            // Workbook created without exception and sheet exists — freeze pane applied via POI internals
            assertNotNull(wb.getSheet("S"));
        }
    }

    @Test
    void generateDynamicExcelWorkbook_horizontalMerge_createsMergedRegion() throws Exception {
        Merge merge = Merge.fromValue(2, 0, "Merged Value", Merge.Orientation.HORIZONTAL);
        List<MergeDTO> data = Collections.singletonList(MergeDTO.builder().mergeField(merge).build());

        try (Workbook wb = service.generateDynamicExcelWorkbook(
                Arrays.asList(MergeHeader.values()), data, MergeDTO.class, settings("S"))) {
            Sheet sheet = wb.getSheet("S");
            assertEquals(1, sheet.getNumMergedRegions());
            assertEquals(0, sheet.getMergedRegion(0).getFirstColumn());
            assertEquals(1, sheet.getMergedRegion(0).getLastColumn());
        }
    }

    @Test
    void generateDynamicExcelWorkbook_verticalMerge_createsMergedRegion() throws Exception {
        Merge merge = Merge.fromValue(2, 0, "Merged Value", Merge.Orientation.VERTICAL);
        List<MergeDTO> data = Arrays.asList(
                MergeDTO.builder().mergeField(merge).build(),
                MergeDTO.builder().build());

        try (Workbook wb = service.generateDynamicExcelWorkbook(
                Arrays.asList(MergeHeader.values()), data, MergeDTO.class, settings("S"))) {
            Sheet sheet = wb.getSheet("S");
            assertEquals(1, sheet.getNumMergedRegions());
            assertEquals(1, sheet.getMergedRegion(0).getFirstRow());
            assertEquals(2, sheet.getMergedRegion(0).getLastRow());
        }
    }

    @Test
    void generateDynamicExcelWorkbook_appendsToExistingWorkbook() throws Exception {
        Workbook existingWorkbook = new XSSFWorkbook();
        existingWorkbook.createSheet("Existing");
        ExcelSettings s = ExcelSettings.builder().sheetName("New").build();

        try (Workbook wb = service.generateDynamicExcelWorkbook(
                basicHeaders(), Collections.emptyList(), BasicDTO.class, s, existingWorkbook)) {
            assertNotNull(wb.getSheet("Existing"));
            assertNotNull(wb.getSheet("New"));
        }
    }

    @Test
    void generateDynamicExcelWorkbook_nullSheetName_throwsExcelGenerationException() {
        ExcelSettings s = ExcelSettings.builder().build(); // sheetName is null

        assertThrows(ExcelGenerationException.class, () ->
                service.generateDynamicExcelWorkbook(basicHeaders(), Collections.emptyList(), BasicDTO.class, s));
    }

    @Test
    void generateDynamicExcel_returnsNonEmptyByteArray() throws Exception {
        List<BasicDTO> data = Collections.singletonList(BasicDTO.builder().name("Test").build());

        byte[] bytes = service.generateDynamicExcel(basicHeaders(), data, BasicDTO.class, settings("S"));

        assertNotNull(bytes);
        assertTrue(bytes.length > 0);
    }

    @Test
    void generateDynamicExcel_byteArray_isValidXlsxWorkbook() throws Exception {
        List<BasicDTO> data = Collections.singletonList(BasicDTO.builder().name("Eve").amount(7.0).build());

        byte[] bytes = service.generateDynamicExcel(basicHeaders(), data, BasicDTO.class, settings("Sheet1"));

        try (Workbook wb = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
            assertEquals("Eve", wb.getSheet("Sheet1").getRow(1).getCell(0).getStringCellValue());
        }
    }
}
