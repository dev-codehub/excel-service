package io.github.devcodehub.excel.lib.service;

import io.github.devcodehub.excel.lib.model.dto.csv.CsvHeaderBase;
import io.github.devcodehub.excel.lib.model.dto.csv.CsvSettings;
import io.github.devcodehub.excel.lib.model.dto.exception.CsvGenerationException;
import io.github.devcodehub.excel.lib.service.csv.CsvServiceImpl;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.nio.charset.StandardCharsets;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

class CsvServiceImplTest {

    private CsvServiceImpl service;
    private CsvSettings defaultSettings;

    @BeforeEach
    void setUp() {
        service = new CsvServiceImpl();
        defaultSettings = CsvSettings.builder().build();
    }

    // ── DTOs ──────────────────────────────────────────────────────────────────

    @Getter @Setter @Builder @NoArgsConstructor @AllArgsConstructor
    static class PersonDTO {
        private String name;
        private Integer age;
    }

    @Getter @Setter @Builder @NoArgsConstructor @AllArgsConstructor
    static class AllTypesDTO {
        private String strField;
        private Integer intField;
        private Long longField;
        private Double doubleField;
        private Float floatField;
        private Boolean boolField;
    }

    static class NoDefaultConstructorDTO {
        private final String value;
        NoDefaultConstructorDTO(String value) { this.value = value; }
    }

    // ── Headers ───────────────────────────────────────────────────────────────

    @AllArgsConstructor
    enum PersonHeader implements CsvHeaderBase {
        NAME("name", "Name"),
        AGE("age", "Age"),
        ;
        private final String field;
        private final String displayName;
        @Override public String getField() { return field; }
        @Override public String getDisplayName() { return displayName; }
    }

    @AllArgsConstructor
    enum AllTypesHeader implements CsvHeaderBase {
        STR("strField", "String"),
        INT("intField", "Integer"),
        LONG("longField", "Long"),
        DOUBLE("doubleField", "Double"),
        FLOAT("floatField", "Float"),
        BOOL("boolField", "Boolean"),
        ;
        private final String field;
        private final String displayName;
        @Override public String getField() { return field; }
        @Override public String getDisplayName() { return displayName; }
    }

    // ── generateCsv ───────────────────────────────────────────────────────────

    @Test
    void generateCsv_producesHeaderRowAndDataRow() throws CsvGenerationException {
        List<PersonDTO> data = Arrays.asList(
                PersonDTO.builder().name("Alice").age(30).build(),
                PersonDTO.builder().name("Bob").age(25).build()
        );

        byte[] result = service.generateCsv(
                Arrays.asList(PersonHeader.values()), data, PersonDTO.class, defaultSettings);

        String csv = new String(result, StandardCharsets.UTF_8);
        assertTrue(csv.contains("Name"));
        assertTrue(csv.contains("Age"));
        assertTrue(csv.contains("Alice"));
        assertTrue(csv.contains("30"));
        assertTrue(csv.contains("Bob"));
        assertTrue(csv.contains("25"));
    }

    @Test
    void generateCsv_emptyData_throwsCsvGenerationException() {
        assertThrows(CsvGenerationException.class, () ->
                service.generateCsv(Arrays.asList(PersonHeader.values()),
                        Collections.emptyList(), PersonDTO.class, defaultSettings));
    }

    @Test
    void generateCsv_customDelimiter_usesDelimiterInOutput() throws CsvGenerationException {
        CsvSettings commaSettings = CsvSettings.builder().delimiter(',').build();
        List<PersonDTO> data = Collections.singletonList(
                PersonDTO.builder().name("Alice").age(30).build());

        byte[] result = service.generateCsv(
                Arrays.asList(PersonHeader.values()), data, PersonDTO.class, commaSettings);

        String csv = new String(result, StandardCharsets.UTF_8);
        assertTrue(csv.contains("Name,Age"));
    }

    // ── readCsv ───────────────────────────────────────────────────────────────

    @Test
    void readCsv_roundTrip_returnsOriginalData() throws CsvGenerationException {
        List<PersonDTO> original = Arrays.asList(
                PersonDTO.builder().name("Alice").age(30).build(),
                PersonDTO.builder().name("Bob").age(25).build()
        );
        List<? extends CsvHeaderBase> headers = Arrays.asList(PersonHeader.values());

        byte[] csv = service.generateCsv(headers, original, PersonDTO.class, defaultSettings);
        List<PersonDTO> result = service.readCsv(csv, headers, PersonDTO.class, defaultSettings);

        assertEquals(2, result.size());
        assertEquals("Alice", result.get(0).getName());
        assertEquals(30, result.get(0).getAge());
        assertEquals("Bob", result.get(1).getName());
        assertEquals(25, result.get(1).getAge());
    }

    @Test
    void readCsv_columnOrderIndependent_mapsCorrectly() throws CsvGenerationException {
        // CSV with columns in reversed order: Age;Name
        String csv = "Age;Name\r\n30;Alice\r\n25;Bob\r\n";
        byte[] bytes = csv.getBytes(StandardCharsets.UTF_8);

        List<PersonDTO> result = service.readCsv(
                bytes, Arrays.asList(PersonHeader.values()), PersonDTO.class, defaultSettings);

        assertEquals(2, result.size());
        assertEquals("Alice", result.get(0).getName());
        assertEquals(30, result.get(0).getAge());
    }

    @Test
    void readCsv_allSupportedTypes_convertsCorrectly() throws CsvGenerationException {
        String csv = "String;Integer;Long;Double;Float;Boolean\r\nhello;42;9876543210;3.14;1.5;true\r\n";
        byte[] bytes = csv.getBytes(StandardCharsets.UTF_8);

        List<AllTypesDTO> result = service.readCsv(
                bytes, Arrays.asList(AllTypesHeader.values()), AllTypesDTO.class, defaultSettings);

        assertEquals(1, result.size());
        AllTypesDTO dto = result.get(0);
        assertEquals("hello", dto.getStrField());
        assertEquals(42, dto.getIntField());
        assertEquals(9876543210L, dto.getLongField());
        assertEquals(3.14, dto.getDoubleField(), 0.001);
        assertEquals(1.5f, dto.getFloatField(), 0.001f);
        assertTrue(dto.getBoolField());
    }

    @Test
    void readCsv_missingNoArgsConstructor_throwsCsvGenerationException() {
        String csv = "Name;Age\r\nAlice;30\r\n";
        byte[] bytes = csv.getBytes(StandardCharsets.UTF_8);

        assertThrows(CsvGenerationException.class, () ->
                service.readCsv(bytes, Arrays.asList(PersonHeader.values()),
                        NoDefaultConstructorDTO.class, defaultSettings));
    }

    @Test
    void readCsv_unknownColumnInFile_isIgnored() throws CsvGenerationException {
        // CSV has an extra column "Email" not in the header enum
        String csv = "Name;Age;Email\r\nAlice;30;alice@example.com\r\n";
        byte[] bytes = csv.getBytes(StandardCharsets.UTF_8);

        List<PersonDTO> result = service.readCsv(
                bytes, Arrays.asList(PersonHeader.values()), PersonDTO.class, defaultSettings);

        assertEquals(1, result.size());
        assertNotNull(result.get(0).getName());
    }

    @Test
    void readCsv_emptyCsv_returnsEmptyList() throws CsvGenerationException {
        String csv = "Name;Age\r\n";
        byte[] bytes = csv.getBytes(StandardCharsets.UTF_8);

        List<PersonDTO> result = service.readCsv(
                bytes, Arrays.asList(PersonHeader.values()), PersonDTO.class, defaultSettings);

        assertTrue(result.isEmpty());
    }
}
