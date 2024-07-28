package com.excel.lib.service;

import com.excel.lib.model.dto.excel.ExcelColor;
import com.excel.lib.model.dto.excel.ExcelCustomStyles;
import com.excel.lib.model.dto.excel.StyleDTO;
import com.excel.lib.model.dto.excel.datatype.DateExcel;
import com.excel.lib.model.dto.excel.datatype.Merge;
import com.excel.lib.model.dto.excel.datatype.StringExcel;
import com.excel.lib.utils.ExcelUtils;
import lombok.Getter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class StyleService {
    public static final List<XSSFCellBorder.BorderSide> BORDER_SIDES = List.of(XSSFCellBorder.BorderSide.TOP, XSSFCellBorder.BorderSide.BOTTOM, XSSFCellBorder.BorderSide.LEFT, XSSFCellBorder.BorderSide.RIGHT);

    @Getter
    private final CellStyle headerCellStyle;
    @Getter
    private final CellStyle dataCellStyle;
    private final Map<Class<?>, CellStyle> stylesMap;


    public StyleService(Workbook workbook, ExcelCustomStyles excelCustomStyles) {
        // Header and data default styles
        headerCellStyle = workbook.createCellStyle();
        dataCellStyle = workbook.createCellStyle();
        setDefaultStyles(workbook, excelCustomStyles);

        stylesMap = new HashMap<>();

        // Merge default styles
        CellStyle mergeCellStyle = workbook.createCellStyle();
        mergeCellStyle.cloneStyleFrom(dataCellStyle);
        stylesMap.put(Merge.class, mergeCellStyle);

        // Date default styles
        CellStyle dateCellStyle = workbook.createCellStyle();
        dateCellStyle.cloneStyleFrom(dataCellStyle);
        stylesMap.put(DateExcel.class, dateCellStyle);

        // String default styles
        CellStyle stringCellStyle = workbook.createCellStyle();
        stringCellStyle.cloneStyleFrom(dataCellStyle);
        stylesMap.put(StringExcel.class, stringCellStyle);
    }

    private void setDefaultStyles(Workbook workbook, ExcelCustomStyles excelCustomStyles) {
        // header default style
        Font font = workbook.createFont();
        font.setFontName(excelCustomStyles.getFontFamily());
        font.setBold(excelCustomStyles.isHeaderBold());
        font.setFontHeightInPoints(excelCustomStyles.getHeaderFontHeight());
        ((XSSFFont) font).setColor(ExcelColor.BLACK.getColor());
        headerCellStyle.setFont(font);
        headerCellStyle.setWrapText(true);
        headerCellStyle.setAlignment(HorizontalAlignment.CENTER);
        headerCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headerCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerCellStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());
        if (excelCustomStyles.isHeaderBorderActive()) {
            headerCellStyle.setBorderTop(BorderStyle.THIN);
            headerCellStyle.setBorderRight(BorderStyle.THIN);
            headerCellStyle.setBorderBottom(BorderStyle.THIN);
            headerCellStyle.setBorderLeft(BorderStyle.THIN);
            BORDER_SIDES.forEach(border -> ((XSSFCellStyle) headerCellStyle).setBorderColor(border,
                    excelCustomStyles.getHeaderBorderColor().getColor()));
        }

        // data default style
        font = workbook.createFont();
        font.setFontName(excelCustomStyles.getFontFamily());
        font.setBold(excelCustomStyles.isDataCellBold());
        font.setFontHeightInPoints(excelCustomStyles.getDataCellFontHeight());
        dataCellStyle.setFont(font);
        dataCellStyle.setWrapText(true);
        dataCellStyle.setAlignment(HorizontalAlignment.CENTER);
        dataCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        if (excelCustomStyles.isDataBorderActive()) {
            dataCellStyle.setBorderTop(BorderStyle.THIN);
            dataCellStyle.setBorderRight(BorderStyle.THIN);
            dataCellStyle.setBorderBottom(BorderStyle.THIN);
            dataCellStyle.setBorderLeft(BorderStyle.THIN);
            BORDER_SIDES.forEach(border -> ((XSSFCellStyle) dataCellStyle).setBorderColor(border, excelCustomStyles.getDataBorderColor().getColor()));
        }
    }

    // TODO: replace by getCellStyle
    public CellStyle getHeaderCellStyle(Workbook workbook, StyleDTO styles) {
        if (styles == null) {
            return headerCellStyle;
        }

        // Create a new cell style or the new styles will not be applied
        CellStyle newCellStyle = workbook.createCellStyle();
        newCellStyle.cloneStyleFrom(headerCellStyle);

        // Get font from cell style
        Font font = ExcelUtils.cloneFont(workbook, workbook.getFontAt(headerCellStyle.getFontIndex()));

        // Apply new styles and keep default values if not updated
        setCellStyle(styles, newCellStyle, font);

        return newCellStyle;
    }

    public CellStyle getCellStyle(Workbook workbook, Class<?> type, StyleDTO styles) {
        CellStyle cellStyle = stylesMap.get(type);
        if (styles == null) {
            return cellStyle;
        }

        // Create a new cell style or the new styles will not be applied
        CellStyle newCellStyle = workbook.createCellStyle();
        newCellStyle.cloneStyleFrom(cellStyle);

        // Get font from cell style
        Font font = ExcelUtils.cloneFont(workbook, workbook.getFontAt(cellStyle.getFontIndex()));

        // Apply new styles and keep default values if not updated
        setCellStyle(styles, newCellStyle, font);

        return newCellStyle;
    }
    /**
     * This method is responsible to apply styles in StyleDTO to an Excel cell
     */
    private void setCellStyle(StyleDTO styles, CellStyle newCellStyle, Font font) {
        // Set alignments
        if (styles.getHorizontalAlignment() != null)
            newCellStyle.setAlignment(styles.getHorizontalAlignment());
        if (styles.getVerticalAlignment() != null)
            newCellStyle.setVerticalAlignment(styles.getVerticalAlignment());

        // Set foreground color
        if (styles.getForegroundColor() != null) {
            newCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            newCellStyle.setFillForegroundColor(styles.getForegroundColor().getColor());
        }

        // Set text color
        if (styles.getTextColor() != null)
            ((XSSFFont) font).setColor(styles.getTextColor().getColor());

        // Set font height
        if (styles.getFontHeight() != null)
            font.setFontHeightInPoints(styles.getFontHeight().shortValue());

        // Set bold
        if (styles.getBold() != null)
            font.setBold(styles.getBold());

        newCellStyle.setFont(font);
    }
}
