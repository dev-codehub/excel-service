package io.github.devcodehub.excel.lib.utils;

import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotSame;
import static org.junit.jupiter.api.Assertions.assertTrue;

class ExcelUtilsTest {

    @Test
    void cloneFont_isNotSameInstance() throws Exception {
        try (Workbook wb = new XSSFWorkbook()) {
            Font original = wb.createFont();

            Font cloned = ExcelUtils.cloneFont(wb, original);

            assertNotSame(original, cloned);
        }
    }

    @Test
    void cloneFont_copiesAllProperties() throws Exception {
        try (Workbook wb = new XSSFWorkbook()) {
            Font original = wb.createFont();
            original.setFontName("Arial");
            original.setFontHeightInPoints((short) 14);
            original.setBold(true);
            original.setItalic(true);
            original.setStrikeout(true);
            original.setUnderline(Font.U_SINGLE);

            Font cloned = ExcelUtils.cloneFont(wb, original);

            assertEquals("Arial", cloned.getFontName());
            assertEquals(14, cloned.getFontHeightInPoints());
            assertTrue(cloned.getBold());
            assertTrue(cloned.getItalic());
            assertTrue(cloned.getStrikeout());
            assertEquals(Font.U_SINGLE, cloned.getUnderline());
        }
    }
}
