package com.excel.lib.service;

import com.excel.lib.model.dto.example.ExampleDTO;
import com.excel.lib.model.dto.example.ExcelExampleHeader;
import com.excel.lib.model.dto.excel.ExcelColor;
import com.excel.lib.model.dto.excel.ExcelSettings;
import com.excel.lib.model.dto.excel.StyleDTO;
import com.excel.lib.model.dto.excel.datatype.StringExcel;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;

public class ExcelServiceImplTest {
    private ExcelService excelService;

    @BeforeEach
    void setUp() {
        excelService = new ExcelServiceImpl();
    }

    @Test
    void testGenerateDynamicExcelWorkbook() throws Exception {
        // Headers
        List<ExcelExampleHeader> headers = List.of(ExcelExampleHeader.values());

        // List of objects
        List<ExampleDTO> exampleDTOList = getExampleDTOList();

        // Excel settings
        ExcelSettings excelSettings = ExcelSettings.builder()
                .sheetName("example")
                .rowOffset(1)
                .colOffset(1)
                .freezePane(ExcelSettings.FreezePane.builder().rowSplit(2).colSplit(2).build())
                .build();

        // Generate excel
        Workbook workbook = excelService.generateDynamicExcelWorkbook(headers, exampleDTOList, ExampleDTO.class, excelSettings, null);

        // Verify sheet
        Sheet sheet = workbook.getSheetAt(0);
        assertNotNull(sheet);

        // Verify cell value
        Row row = sheet.getRow(1);
        Cell cell = row.getCell(1);
        assertNotNull(row);
        assertNotNull(cell);
        assertEquals("Display Column 1", cell.getStringCellValue());

        // Verify cell style
        CellStyle style = cell.getCellStyle();
        assertEquals(ExcelColor.DARK_MINT.getColor(), style.getFillForegroundColorColor());

        // Verify cell font
        Font font = workbook.getFontAt(style.getFontIndex());
        assertEquals("Calibri", font.getFontName());
        assertEquals(12, font.getFontHeightInPoints());
        assertFalse(font.getBold());
        assertEquals(ExcelColor.WHITE.getColor().getIndex(), font.getColor());
    }

    private List<ExampleDTO> getExampleDTOList() {
        StyleDTO name1Style = StyleDTO.builder()
                .foregroundColor(ExcelColor.LIGHT_GREY)
                .bold(true)
                .build();
        ExampleDTO example1 = ExampleDTO.builder()
                .value1(StringExcel.fromValue("Value1", name1Style))
                .value2("value2")
                .value3("value3")
                .build();
        ExampleDTO example2 = ExampleDTO.builder()
                .value1(StringExcel.fromValue("Value1", name1Style))
                .value2("value2")
                .value3("value3")
                .build();
        ExampleDTO example3 = ExampleDTO.builder()
                .value1(StringExcel.fromValue("Value1", name1Style))
                .value2("value2")
                .value3("value3")
                .build();
        return List.of(example1, example2, example3);
    }
}
