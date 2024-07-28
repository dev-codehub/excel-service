package com.excel.lib.utils;

import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;

public class ExcelUtils {
    private ExcelUtils() {
        // prevent initialization
    }

    public static Font cloneFont(Workbook workbook, Font font) {
        Font newFont = workbook.createFont();

        newFont.setFontName(font.getFontName());
        newFont.setFontHeightInPoints(font.getFontHeightInPoints());
        newFont.setBold(font.getBold());
        newFont.setItalic(font.getItalic());
        newFont.setStrikeout(font.getStrikeout());
        newFont.setColor(font.getColor());
        newFont.setTypeOffset(font.getTypeOffset());
        newFont.setUnderline(font.getUnderline());
        newFont.setCharSet(font.getCharSet());

        return newFont;
    }
}
