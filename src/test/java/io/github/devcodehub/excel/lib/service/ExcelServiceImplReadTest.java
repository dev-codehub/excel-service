package io.github.devcodehub.excel.lib.service;

import io.github.devcodehub.excel.lib.model.dto.excel.ExcelHeaderBase;
import io.github.devcodehub.excel.lib.model.dto.excel.ExcelReadSettings;
import io.github.devcodehub.excel.lib.model.dto.excel.ExcelSettings;
import io.github.devcodehub.excel.lib.model.dto.excel.StyleDTO;
import io.github.devcodehub.excel.lib.model.dto.excel.datatype.DateExcel;
import io.github.devcodehub.excel.lib.model.dto.excel.datatype.Number;
import io.github.devcodehub.excel.lib.model.dto.excel.datatype.StringExcel;
import io.github.devcodehub.excel.lib.model.dto.exception.ExcelGenerationException;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayOutputStream;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Collections;
import java.util.Date;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

class ExcelServiceImplReadTest {

    private ExcelServiceImpl service;

    @BeforeEach
    void setUp() {
        service = new ExcelServiceImpl();
    }

    // ── DTOs ──────────────────────────────────────────────────────────────────

    @Getter @Setter @Builder @NoArgsConstructor @AllArgsConstructor
    static class PersonDTO {
        private String name;
        private Double amount;
    }

    @Getter @Setter @Builder @NoArgsConstructor @AllArgsConstructor
    static class TypedDTO {
        private String strField;
        private Integer intField;
        private Long longField;
        private Boolean boolField;
        private StringExcel stringExcelField;
        private Number numberField;
    }

    static class NoDefaultConstructorDTO {
        NoDefaultConstructorDTO(String unused) {}
    }

    // ── Headers ───────────────────────────────────────────────────────────────

    @AllArgsConstructor
    enum PersonHeader implements ExcelHeaderBase {
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
    enum TypedHeader implements ExcelHeaderBase {
        STR_FIELD("strField", "StrField", null),
        INT_FIELD("intField", "IntField", null),
        LONG_FIELD("longField", "LongField", null),
        BOOL_FIELD("boolField", "BoolField", null),
        STRING_EXCEL_FIELD("stringExcelField", "StringExcelField", null),
        NUMBER_FIELD("numberField", "NumberField", null);
        private final String field;
        private final String displayName;
        private final StyleDTO styles;
        @Override public String getField() { return field; }
        @Override public String getDisplayName() { return displayName; }
        @Override public StyleDTO getStyles() { return styles; }
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    private static ExcelSettings writeSettings(String sheetName) {
        return ExcelSettings.builder().sheetName(sheetName).build();
    }

    private static ExcelReadSettings readSettings(String sheetName) {
        return ExcelReadSettings.builder().sheetName(sheetName).build();
    }

    private static Date buildDate(int year, int month, int day) {
        Calendar cal = Calendar.getInstance();
        cal.set(year, month, day, 0, 0, 0);
        cal.set(Calendar.MILLISECOND, 0);
        return cal.getTime();
    }

    /** Write a workbook to bytes and return them. */
    private byte[] writeToBytes(Workbook wb) throws Exception {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        wb.write(out);
        wb.close();
        return out.toByteArray();
    }

    // ── Tests ─────────────────────────────────────────────────────────────────

    @Test
    void readDynamicExcel_nullSheetName_throwsExcelGenerationException() {
        ExcelReadSettings s = ExcelReadSettings.builder().build(); // sheetName null

        try (Workbook wb = new XSSFWorkbook()) {
            assertThrows(ExcelGenerationException.class,
                    () -> service.readDynamicExcel(wb, Arrays.asList(PersonHeader.values()), PersonDTO.class, s));
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    @Test
    void readDynamicExcel_sheetNotFound_throwsExcelGenerationException() {
        ExcelReadSettings s = readSettings("DoesNotExist");

        try (Workbook wb = new XSSFWorkbook()) {
            wb.createSheet("OtherSheet");
            assertThrows(ExcelGenerationException.class,
                    () -> service.readDynamicExcel(wb, Arrays.asList(PersonHeader.values()), PersonDTO.class, s));
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    @Test
    void readDynamicExcel_noDefaultConstructor_throwsExcelGenerationException() throws Exception {
        byte[] bytes = service.generateDynamicExcel(
                Arrays.asList(PersonHeader.values()),
                Collections.singletonList(PersonDTO.builder().name("X").build()),
                PersonDTO.class, writeSettings("S"));

        assertThrows(ExcelGenerationException.class,
                () -> service.readDynamicExcel(bytes, Arrays.asList(PersonHeader.values()),
                        NoDefaultConstructorDTO.class, readSettings("S")));
    }

    @Test
    void readDynamicExcel_noMatchingColumns_returnsEmptyList() throws Exception {
        @AllArgsConstructor
        class UnknownHeader implements ExcelHeaderBase {
            @Override public String getField() { return "name"; }
            @Override public String getDisplayName() { return "Completely Unknown Column"; }
            @Override public StyleDTO getStyles() { return null; }
        }

        byte[] bytes = service.generateDynamicExcel(
                Arrays.asList(PersonHeader.values()),
                Collections.singletonList(PersonDTO.builder().name("Alice").build()),
                PersonDTO.class, writeSettings("S"));

        List<PersonDTO> result = service.readDynamicExcel(bytes,
                Collections.singletonList(new UnknownHeader()), PersonDTO.class, readSettings("S"));

        assertTrue(result.isEmpty());
    }

    @Test
    void readDynamicExcel_roundTrip_stringFields() throws Exception {
        // Plain Double fields are written as STRING cells; round-trip uses string-only data
        List<PersonDTO> original = Arrays.asList(
                PersonDTO.builder().name("Alice").build(),
                PersonDTO.builder().name("Bob").build());

        byte[] bytes = service.generateDynamicExcel(
                Arrays.asList(PersonHeader.values()), original, PersonDTO.class, writeSettings("S"));

        List<PersonDTO> result = service.readDynamicExcel(bytes,
                Arrays.asList(PersonHeader.values()), PersonDTO.class, readSettings("S"));

        assertEquals(2, result.size());
        assertEquals("Alice", result.get(0).getName());
        assertEquals("Bob", result.get(1).getName());
    }

    @Test
    void readDynamicExcel_withRowOffset_readsCorrectly() throws Exception {
        List<PersonDTO> original = Collections.singletonList(
                PersonDTO.builder().name("Charlie").build());

        ExcelSettings writeS = ExcelSettings.builder().sheetName("S").rowOffset(3).build();
        byte[] bytes = service.generateDynamicExcel(
                Arrays.asList(PersonHeader.values()), original, PersonDTO.class, writeS);

        ExcelReadSettings readS = ExcelReadSettings.builder().sheetName("S").rowOffset(3).build();
        List<PersonDTO> result = service.readDynamicExcel(bytes,
                Arrays.asList(PersonHeader.values()), PersonDTO.class, readS);

        assertEquals(1, result.size());
        assertEquals("Charlie", result.get(0).getName());
    }

    @Test
    void readDynamicExcel_withColOffset_readsCorrectly() throws Exception {
        List<PersonDTO> original = Collections.singletonList(
                PersonDTO.builder().name("Dana").build());

        ExcelSettings writeS = ExcelSettings.builder().sheetName("S").colOffset(2).build();
        byte[] bytes = service.generateDynamicExcel(
                Arrays.asList(PersonHeader.values()), original, PersonDTO.class, writeS);

        ExcelReadSettings readS = ExcelReadSettings.builder().sheetName("S").colOffset(2).build();
        List<PersonDTO> result = service.readDynamicExcel(bytes,
                Arrays.asList(PersonHeader.values()), PersonDTO.class, readS);

        assertEquals(1, result.size());
        assertEquals("Dana", result.get(0).getName());
    }

    @Test
    void readDynamicExcel_searchKeys_findsTableByKeys() throws Exception {
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("S");

        // Row 0: preamble
        Row preamble = sheet.createRow(0);
        preamble.createCell(0).setCellValue("Report Title");

        // Row 1: empty
        sheet.createRow(1);

        // Row 2: header
        Row header = sheet.createRow(2);
        header.createCell(0).setCellValue("Name");
        header.createCell(1).setCellValue("Amount");

        // Row 3: data
        Row data = sheet.createRow(3);
        data.createCell(0).setCellValue("Eve");
        data.createCell(1).setCellValue(99.0);

        byte[] bytes = writeToBytes(wb);

        ExcelReadSettings readS = ExcelReadSettings.builder()
                .sheetName("S")
                .searchKeys(Arrays.asList("Name", "Amount"))
                .build();

        List<PersonDTO> result = service.readDynamicExcel(bytes,
                Arrays.asList(PersonHeader.values()), PersonDTO.class, readS);

        assertEquals(1, result.size());
        assertEquals("Eve", result.get(0).getName());
        assertEquals(99.0, result.get(0).getAmount(), 0.001);
    }

    @Test
    void readDynamicExcel_searchKeysNotFound_throwsExcelGenerationException() throws Exception {
        byte[] bytes = service.generateDynamicExcel(
                Arrays.asList(PersonHeader.values()),
                Collections.singletonList(PersonDTO.builder().name("X").build()),
                PersonDTO.class, writeSettings("S"));

        ExcelReadSettings readS = ExcelReadSettings.builder()
                .sheetName("S")
                .searchKeys(Arrays.asList("NonExistent", "Keys"))
                .build();

        assertThrows(ExcelGenerationException.class,
                () -> service.readDynamicExcel(bytes, Arrays.asList(PersonHeader.values()),
                        PersonDTO.class, readS));
    }

    @Test
    void readDynamicExcel_searchKeys_maxScanRows_limitsSearch() throws Exception {
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("S");

        // Header at row 5 — beyond the scan limit of 3
        for (int i = 0; i < 5; i++) sheet.createRow(i);
        Row header = sheet.createRow(5);
        header.createCell(0).setCellValue("Name");
        header.createCell(1).setCellValue("Amount");

        byte[] bytes = writeToBytes(wb);

        ExcelReadSettings readS = ExcelReadSettings.builder()
                .sheetName("S")
                .searchKeys(Arrays.asList("Name", "Amount"))
                .maxScanRows(3)
                .build();

        assertThrows(ExcelGenerationException.class,
                () -> service.readDynamicExcel(bytes, Arrays.asList(PersonHeader.values()),
                        PersonDTO.class, readS));
    }

    @Test
    void readDynamicExcel_emptyRows_areSkipped() throws Exception {
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("S");

        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("Name");
        header.createCell(1).setCellValue("Amount");

        Row data1 = sheet.createRow(1);
        data1.createCell(0).setCellValue("Alice");
        data1.createCell(1).setCellValue(1.0);

        // Row 2: empty (skip it)
        sheet.createRow(2);

        Row data3 = sheet.createRow(3);
        data3.createCell(0).setCellValue("Bob");
        data3.createCell(1).setCellValue(2.0);

        byte[] bytes = writeToBytes(wb);

        List<PersonDTO> result = service.readDynamicExcel(bytes,
                Arrays.asList(PersonHeader.values()), PersonDTO.class, readSettings("S"));

        assertEquals(2, result.size());
        assertEquals("Alice", result.get(0).getName());
        assertEquals("Bob", result.get(1).getName());
    }

    @Test
    void readDynamicExcel_nullCell_fieldLeftNull() throws Exception {
        List<PersonDTO> original = Collections.singletonList(
                PersonDTO.builder().name("Fiona").build()); // amount is null

        byte[] bytes = service.generateDynamicExcel(
                Arrays.asList(PersonHeader.values()), original, PersonDTO.class, writeSettings("S"));

        List<PersonDTO> result = service.readDynamicExcel(bytes,
                Arrays.asList(PersonHeader.values()), PersonDTO.class, readSettings("S"));

        assertEquals(1, result.size());
        assertEquals("Fiona", result.get(0).getName());
        assertEquals(null, result.get(0).getAmount());
    }

    @Test
    void readDynamicExcel_booleanCell_mappedToBoolean() throws Exception {
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("S");

        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("StrField");
        header.createCell(1).setCellValue("IntField");
        header.createCell(2).setCellValue("LongField");
        header.createCell(3).setCellValue("BoolField");
        header.createCell(4).setCellValue("StringExcelField");
        header.createCell(5).setCellValue("NumberField");

        Row data = sheet.createRow(1);
        data.createCell(3).setCellValue(true);

        byte[] bytes = writeToBytes(wb);

        List<TypedDTO> result = service.readDynamicExcel(bytes,
                Arrays.asList(TypedHeader.values()), TypedDTO.class, readSettings("S"));

        assertEquals(1, result.size());
        assertEquals(Boolean.TRUE, result.get(0).getBoolField());
    }

    @Test
    void readDynamicExcel_integerCell_mappedToInteger() throws Exception {
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("S");

        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("StrField");
        header.createCell(1).setCellValue("IntField");
        header.createCell(2).setCellValue("LongField");
        header.createCell(3).setCellValue("BoolField");
        header.createCell(4).setCellValue("StringExcelField");
        header.createCell(5).setCellValue("NumberField");

        Row data = sheet.createRow(1);
        data.createCell(1).setCellValue(42.0);

        byte[] bytes = writeToBytes(wb);

        List<TypedDTO> result = service.readDynamicExcel(bytes,
                Arrays.asList(TypedHeader.values()), TypedDTO.class, readSettings("S"));

        assertEquals(1, result.size());
        assertEquals(Integer.valueOf(42), result.get(0).getIntField());
    }

    @Test
    void readDynamicExcel_longCell_mappedToLong() throws Exception {
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("S");

        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("StrField");
        header.createCell(1).setCellValue("IntField");
        header.createCell(2).setCellValue("LongField");
        header.createCell(3).setCellValue("BoolField");
        header.createCell(4).setCellValue("StringExcelField");
        header.createCell(5).setCellValue("NumberField");

        Row data = sheet.createRow(1);
        data.createCell(2).setCellValue(1234567890.0);

        byte[] bytes = writeToBytes(wb);

        List<TypedDTO> result = service.readDynamicExcel(bytes,
                Arrays.asList(TypedHeader.values()), TypedDTO.class, readSettings("S"));

        assertEquals(1, result.size());
        assertEquals(Long.valueOf(1234567890L), result.get(0).getLongField());
    }

    @Test
    void readDynamicExcel_stringExcelField_mappedToStringExcel() throws Exception {
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("S");

        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("StrField");
        header.createCell(1).setCellValue("IntField");
        header.createCell(2).setCellValue("LongField");
        header.createCell(3).setCellValue("BoolField");
        header.createCell(4).setCellValue("StringExcelField");
        header.createCell(5).setCellValue("NumberField");

        Row data = sheet.createRow(1);
        data.createCell(4).setCellValue("HelloExcel");

        byte[] bytes = writeToBytes(wb);

        List<TypedDTO> result = service.readDynamicExcel(bytes,
                Arrays.asList(TypedHeader.values()), TypedDTO.class, readSettings("S"));

        assertEquals(1, result.size());
        assertNotNull(result.get(0).getStringExcelField());
        assertEquals("HelloExcel", result.get(0).getStringExcelField().getValue());
    }

    @Test
    void readDynamicExcel_numberField_mappedToNumber() throws Exception {
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("S");

        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("StrField");
        header.createCell(1).setCellValue("IntField");
        header.createCell(2).setCellValue("LongField");
        header.createCell(3).setCellValue("BoolField");
        header.createCell(4).setCellValue("StringExcelField");
        header.createCell(5).setCellValue("NumberField");

        Row data = sheet.createRow(1);
        data.createCell(5).setCellValue(3.14);

        byte[] bytes = writeToBytes(wb);

        List<TypedDTO> result = service.readDynamicExcel(bytes,
                Arrays.asList(TypedHeader.values()), TypedDTO.class, readSettings("S"));

        assertEquals(1, result.size());
        assertNotNull(result.get(0).getNumberField());
        assertEquals(3.14, result.get(0).getNumberField().getValue(), 0.001);
    }

    @Test
    void readDynamicExcel_formulaCell_evaluated() throws Exception {
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("S");

        Row header = sheet.createRow(0);
        header.createCell(0).setCellValue("Name");
        header.createCell(1).setCellValue("Amount");

        Row data = sheet.createRow(1);
        data.createCell(0).setCellValue("George");
        data.createCell(1).setCellFormula("1+2");

        byte[] bytes = writeToBytes(wb);

        List<PersonDTO> result = service.readDynamicExcel(bytes,
                Arrays.asList(PersonHeader.values()), PersonDTO.class, readSettings("S"));

        assertEquals(1, result.size());
        assertEquals("George", result.get(0).getName());
        assertEquals(3.0, result.get(0).getAmount(), 0.001);
    }

    @Test
    void readDynamicExcel_fromByteArray_returnsValidList() throws Exception {
        List<PersonDTO> original = Collections.singletonList(
                PersonDTO.builder().name("Henry").build());

        byte[] bytes = service.generateDynamicExcel(
                Arrays.asList(PersonHeader.values()), original, PersonDTO.class, writeSettings("S"));

        List<PersonDTO> result = service.readDynamicExcel(bytes,
                Arrays.asList(PersonHeader.values()), PersonDTO.class, readSettings("S"));

        assertEquals(1, result.size());
        assertEquals("Henry", result.get(0).getName());
    }

    @Test
    void readDynamicExcel_multipleRows_allRead() throws Exception {
        List<PersonDTO> original = Arrays.asList(
                PersonDTO.builder().name("Iris").build(),
                PersonDTO.builder().name("Jack").build(),
                PersonDTO.builder().name("Kate").build());

        byte[] bytes = service.generateDynamicExcel(
                Arrays.asList(PersonHeader.values()), original, PersonDTO.class, writeSettings("S"));

        List<PersonDTO> result = service.readDynamicExcel(bytes,
                Arrays.asList(PersonHeader.values()), PersonDTO.class, readSettings("S"));

        assertEquals(3, result.size());
        assertEquals("Iris", result.get(0).getName());
        assertEquals("Jack", result.get(1).getName());
        assertEquals("Kate", result.get(2).getName());
    }
}
