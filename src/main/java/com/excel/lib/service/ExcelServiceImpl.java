package com.excel.lib.service;

import com.excel.lib.model.dto.excel.ExcelHeaderBase;
import com.excel.lib.model.dto.excel.ExcelSettings;
import com.excel.lib.model.dto.excel.StyleDTO;
import com.excel.lib.model.dto.excel.datatype.DateExcel;
import com.excel.lib.model.dto.excel.datatype.Merge;
import com.excel.lib.model.dto.excel.datatype.Number;
import com.excel.lib.model.dto.excel.datatype.StringExcel;
import com.excel.lib.model.dto.exception.ResponseCode;
import com.excel.lib.utils.DateUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.context.annotation.Lazy;
import org.springframework.stereotype.Service;
import org.springframework.util.StringUtils;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.Date;
import java.util.List;

@Lazy
@Service
public class ExcelServiceImpl implements ExcelService {
    private static final int POI_DEFAULT_UNIT = 256;
    private static final int DEFAULT_COLUMN_WIDTH = 10 * POI_DEFAULT_UNIT;
    private static final int MAX_COLUMN_WIDTH = 255 * POI_DEFAULT_UNIT;
    private static final int DEFAULT_COLUMN_MARGIN = 5 * POI_DEFAULT_UNIT;
    private final Logger log = LoggerFactory.getLogger(this.getClass());
    private StyleService styleService;

    @Override
    public byte[] generateDynamicExcel(List<? extends ExcelHeaderBase> headers, List<?> data, Class<?> dataClass,
                                       ExcelSettings excelSettings) throws Exception {
        return generateDynamicExcel(headers, data, dataClass, excelSettings, null);
    }

    @Override
    public byte[] generateDynamicExcel(List<? extends ExcelHeaderBase> headers, List<?> data, Class<?> dataClass,
                                       ExcelSettings excelSettings, Workbook workbook) throws Exception {
        try (Workbook generatedWorkbook = generateDynamicExcelWorkbook(headers, data, dataClass, excelSettings, workbook)) {
            if (generatedWorkbook != null) {
                // Convert workbook into a byte array
                ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
                generatedWorkbook.write(outputStream);
                return outputStream.toByteArray();
            }
        } catch (IOException e) {
            log.error("Error generating dynamic excel!", e);
        }

        throw new Exception(ResponseCode.ERROR_GENERATING_DYNAMIC_EXCEL.getMessage());
    }

    @Override
    public Workbook generateDynamicExcelWorkbook(List<? extends ExcelHeaderBase> headers, List<?> data,
                                                 Class<?> dataClass, ExcelSettings excelSettings) throws Exception {
        return generateDynamicExcelWorkbook(headers, data, dataClass, excelSettings, null);
    }

    @Override
    public Workbook generateDynamicExcelWorkbook(List<? extends ExcelHeaderBase> headers, List<?> data,
                                                 Class<?> dataClass, ExcelSettings excelSettings,
                                                 Workbook workbook) throws Exception {
        String sheetName = excelSettings.getSheetName();
        log.info("Generating dynamic excel - Sheet name [{}]", sheetName);

        if (sheetName == null) {
            throw new Exception("Sheet name must not be null. It must be defined in ExcelSettings.");
        }

        try {
            if (workbook == null)
                workbook = new XSSFWorkbook();

            Sheet sheet = workbook.getSheet(sheetName);
            if (sheet == null)
                sheet = workbook.createSheet(sheetName);

            // This new instance of the service is necessary because we need to send the created workbook to get new styles from it
            styleService = new StyleService(workbook, excelSettings.getExcelCustomStyles());

            // Create data rows
            int rowOffset = excelSettings.getRowOffset();
            int colOffset = excelSettings.getColOffset();
            int rowIndex = rowOffset + 1;
            for (Object dto : data) {
                Row dataRow = sheet.createRow(rowIndex++);
                int colIndex = colOffset;

                for (ExcelHeaderBase header : headers) {
                    Field field = dataClass.getDeclaredField(header.getField());
                    field.setAccessible(true); // This is required to access private fields, otherwise it will throw IllegalAccessException
                    Object value = field.get(dto);

                    if (value != null) {
                        Cell cell = dataRow.createCell(colIndex);
                        setObjectCellValue(cell, value, sheet, rowIndex - 1, colIndex, workbook);
                    }
                    if (header.getDisplayName() != null)
                        colIndex++;
                }
            }

            // Auto size columns width
            for (int colNum = colOffset; colNum < (headers.size() + colOffset); colNum++) {
                sheet.autoSizeColumn(colNum);

                // Check if this column has minimum width defined on this specific header Enum
                StyleDTO styles = headers.get(colNum - colOffset).getStyles();
                int defaultColumnWidth = styles != null && styles.getMinWidth() != null ? styles.getMinWidth() * POI_DEFAULT_UNIT :
                        DEFAULT_COLUMN_WIDTH;

                // Force width (apache poi default unit of a character is 256)
                sheet.setColumnWidth(colNum, Math.min(styles != null && styles.getMaxWidth() != null ?
                                Math.min(MAX_COLUMN_WIDTH, (styles.getMaxWidth() * POI_DEFAULT_UNIT)) : MAX_COLUMN_WIDTH,
                        DEFAULT_COLUMN_MARGIN + Math.max(defaultColumnWidth, sheet.getColumnWidth(colNum))));
            }

            // Create header row.
            // Note: The headers are created after data rows, so that the auto columns size doesn't take header's width into account
            Row headerRow = sheet.createRow(rowOffset);
            int colIndex = colOffset;
            for (ExcelHeaderBase header : headers) {
                String headerDisplayName = header.getDisplayName();
                if (headerDisplayName != null) {
                    Cell cell = headerRow.createCell(colIndex++);
                    cell.setCellValue(headerDisplayName);

                    StyleDTO styles = header.getStyles();
                    if (styles != null) {
                        CellStyle customHeaderCellStyle = styleService.getHeaderCellStyle(workbook, styles);
                        cell.setCellStyle(customHeaderCellStyle);
                    } else {
                        cell.setCellStyle(styleService.getHeaderCellStyle());
                    }
                }
            }

            // Set headers height
            if (excelSettings.getExcelCustomStyles().getHeadersHeight() != null)
                headerRow.setHeightInPoints(excelSettings.getExcelCustomStyles().getHeadersHeight().floatValue());

            // Set header filters
            if (excelSettings.isHeaderFilterActive())
                sheet.setAutoFilter(new CellRangeAddress(rowOffset, rowOffset, colOffset, headers.size()));

            // Set freeze pane
            ExcelSettings.FreezePane freezePane = excelSettings.getFreezePane();
            if (freezePane != null)
                sheet.createFreezePane(freezePane.getColSplit(), freezePane.getRowSplit(), freezePane.getLeftmostColumn(), freezePane.getTopRow());

            return workbook;
        } catch (ReflectiveOperationException e) {
            log.error("Error generating dynamic excel sheet! Sheet name [{}]", sheetName, e);
        }

        throw new Exception(ResponseCode.ERROR_GENERATING_DYNAMIC_EXCEL.getMessage());
    }


    private void setObjectCellValue(Cell cell, Object value, Sheet sheet, int rowIndex, int colIndex, Workbook workbook) {
        // Default style, this can change depending on the column type bellow
        cell.setCellStyle(styleService.getDataCellStyle());

        if (value instanceof Date date) {
            cell.setCellValue(DateUtils.getDateAsString(date));
        } else if (value instanceof DateExcel dateExcel) {
            cell.setCellValue(dateExcel.getValue());
            cell.setCellStyle(styleService.getCellStyle(workbook, DateExcel.class, dateExcel.getStyles()));
        } else if (value instanceof StringExcel stringExcel) {
            cell.setCellValue(stringExcel.getValue());
            cell.setCellStyle(styleService.getCellStyle(workbook, StringExcel.class, stringExcel.getStyles()));
        } else if (value instanceof Number number) {
            cell.setCellValue(number.getValue());
            cell.setCellStyle(styleService.getCellStyle(workbook, Number.class, number.getStyles()));
        } else if (value instanceof Boolean bool) {
            cell.setCellValue(bool);
        } else if (value instanceof Merge merge) {
            String cellValue = merge.getValue();
            if (StringUtils.hasLength(cellValue)) {
                if (merge.getRange() > 1) {
                    switch (merge.getOrientation()) {
                        case VERTICAL:
                            sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex + merge.getRange() - 1, colIndex + merge.getOffset(),
                                    colIndex + merge.getOffset()));
                            break;
                        case HORIZONTAL:
                            sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, colIndex + merge.getOffset(),
                                    colIndex + merge.getOffset() + merge.getRange() - 1));
                            break;
                        default:
                            break;
                    }
                }

                cell = sheet.getRow(rowIndex).createCell(colIndex + merge.getOffset());
                cell.setCellValue(cellValue);
                cell.setCellStyle(styleService.getCellStyle(workbook, Merge.class, merge.getStyles()));
            }
        } else {
            cell.setCellValue(value == null ? "" : value.toString());
        }
    }
}
