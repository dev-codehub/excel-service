package io.github.devcodehub.excel.lib.service;

import io.github.devcodehub.excel.lib.model.dto.excel.ExcelHeaderBase;
import io.github.devcodehub.excel.lib.model.dto.excel.ExcelReadSettings;
import io.github.devcodehub.excel.lib.model.dto.excel.ExcelSettings;
import io.github.devcodehub.excel.lib.model.dto.excel.StyleDTO;
import io.github.devcodehub.excel.lib.model.dto.excel.datatype.DateExcel;
import io.github.devcodehub.excel.lib.model.dto.excel.datatype.Merge;
import io.github.devcodehub.excel.lib.model.dto.excel.datatype.Number;
import io.github.devcodehub.excel.lib.model.dto.excel.datatype.StringExcel;
import io.github.devcodehub.excel.lib.model.dto.exception.ExcelGenerationException;
import io.github.devcodehub.excel.lib.model.dto.exception.ResponseCode;
import io.github.devcodehub.excel.lib.utils.DateUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.context.annotation.Lazy;
import org.springframework.stereotype.Service;
import org.springframework.util.StringUtils;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

@Lazy
@Service
public class ExcelServiceImpl implements ExcelService {
    private static final int POI_DEFAULT_UNIT = 256;
    private static final int DEFAULT_COLUMN_WIDTH = 10 * POI_DEFAULT_UNIT;
    private static final int MAX_COLUMN_WIDTH = 255 * POI_DEFAULT_UNIT;
    private static final int DEFAULT_COLUMN_MARGIN = 5 * POI_DEFAULT_UNIT;

    private final Logger log = LoggerFactory.getLogger(this.getClass());

    // ── Write ─────────────────────────────────────────────────────────────────

    @Override
    public byte[] generateDynamicExcel(List<? extends ExcelHeaderBase> headers, List<?> data, Class<?> dataClass,
                                       ExcelSettings excelSettings) throws ExcelGenerationException {
        return generateDynamicExcel(headers, data, dataClass, excelSettings, null);
    }

    @Override
    public byte[] generateDynamicExcel(List<? extends ExcelHeaderBase> headers, List<?> data, Class<?> dataClass,
                                       ExcelSettings excelSettings, Workbook workbook) throws ExcelGenerationException {
        try (Workbook generatedWorkbook = generateDynamicExcelWorkbook(headers, data, dataClass, excelSettings, workbook);
             ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
            generatedWorkbook.write(outputStream);
            return outputStream.toByteArray();
        } catch (ExcelGenerationException e) {
            throw e;
        } catch (IOException e) {
            log.error("Error writing excel to output stream!", e);
            throw new ExcelGenerationException(ResponseCode.ERROR_GENERATING_DYNAMIC_EXCEL.getMessage(), e);
        }
    }

    @Override
    public Workbook generateDynamicExcelWorkbook(List<? extends ExcelHeaderBase> headers, List<?> data,
                                                 Class<?> dataClass, ExcelSettings excelSettings) throws ExcelGenerationException {
        return generateDynamicExcelWorkbook(headers, data, dataClass, excelSettings, null);
    }

    @Override
    public Workbook generateDynamicExcelWorkbook(List<? extends ExcelHeaderBase> headers, List<?> data,
                                                 Class<?> dataClass, ExcelSettings excelSettings,
                                                 Workbook workbook) throws ExcelGenerationException {
        String sheetName = excelSettings.getSheetName();
        log.info("Generating dynamic excel - Sheet name [{}]", sheetName);

        if (sheetName == null) {
            throw new ExcelGenerationException("Sheet name must not be null. It must be defined in ExcelSettings.");
        }

        try {
            if (workbook == null)
                workbook = new XSSFWorkbook();

            Sheet sheet = workbook.getSheet(sheetName);
            if (sheet == null)
                sheet = workbook.createSheet(sheetName);

            // StyleService is instantiated per call because POI CellStyle objects are bound to a specific Workbook instance
            StyleService styleService = new StyleService(workbook, excelSettings.getExcelCustomStyles());

            int rowOffset = excelSettings.getRowOffset();
            int colOffset = excelSettings.getColOffset();

            // Create data rows
            int rowIndex = rowOffset + 1;
            for (Object dto : data) {
                Row dataRow = sheet.createRow(rowIndex++);
                int colIndex = colOffset;

                for (ExcelHeaderBase header : headers) {
                    Field field = dataClass.getDeclaredField(header.getField());
                    field.setAccessible(true);
                    Object value = field.get(dto);

                    if (value != null) {
                        Cell cell = dataRow.createCell(colIndex);
                        setObjectCellValue(cell, value, sheet, rowIndex - 1, colIndex, workbook, styleService);
                    }
                    if (header.getDisplayName() != null)
                        colIndex++;
                }
            }

            // Auto-size columns.
            // Note: Headers are created after data rows so that auto-sizing is driven by data width, not header width.
            for (int colNum = colOffset; colNum < headers.size() + colOffset; colNum++) {
                sheet.autoSizeColumn(colNum);
                StyleDTO styles = headers.get(colNum - colOffset).getStyles();
                sheet.setColumnWidth(colNum, calculateColumnWidth(sheet.getColumnWidth(colNum), styles));
            }

            // Create header row
            Row headerRow = sheet.createRow(rowOffset);
            int colIndex = colOffset;
            for (ExcelHeaderBase header : headers) {
                String headerDisplayName = header.getDisplayName();
                if (headerDisplayName != null) {
                    Cell cell = headerRow.createCell(colIndex++);
                    cell.setCellValue(headerDisplayName);

                    StyleDTO styles = header.getStyles();
                    CellStyle headerStyle = styles != null
                            ? styleService.getHeaderCellStyle(workbook, styles)
                            : styleService.getHeaderCellStyle();
                    cell.setCellStyle(headerStyle);
                }
            }

            if (excelSettings.getExcelCustomStyles().getHeadersHeight() != null)
                headerRow.setHeightInPoints(excelSettings.getExcelCustomStyles().getHeadersHeight().floatValue());

            if (excelSettings.isHeaderFilterActive())
                sheet.setAutoFilter(new CellRangeAddress(rowOffset, rowOffset, colOffset, colOffset + headers.size() - 1));

            ExcelSettings.FreezePane freezePane = excelSettings.getFreezePane();
            if (freezePane != null)
                sheet.createFreezePane(freezePane.getColSplit(), freezePane.getRowSplit(), freezePane.getLeftmostColumn(), freezePane.getTopRow());

            return workbook;
        } catch (ReflectiveOperationException e) {
            log.error("Error generating dynamic excel sheet! Sheet name [{}]", sheetName, e);
            throw new ExcelGenerationException(ResponseCode.ERROR_GENERATING_DYNAMIC_EXCEL.getMessage(), e);
        }
    }

    private int calculateColumnWidth(int autoSizedWidth, StyleDTO styles) {
        int minWidth = styles != null && styles.getMinWidth() != null
                ? styles.getMinWidth() * POI_DEFAULT_UNIT
                : DEFAULT_COLUMN_WIDTH;
        int maxWidth = styles != null && styles.getMaxWidth() != null
                ? Math.min(MAX_COLUMN_WIDTH, styles.getMaxWidth() * POI_DEFAULT_UNIT)
                : MAX_COLUMN_WIDTH;
        return Math.min(maxWidth, DEFAULT_COLUMN_MARGIN + Math.max(minWidth, autoSizedWidth));
    }

    private void setObjectCellValue(Cell cell, Object value, Sheet sheet, int rowIndex, int colIndex,
                                    Workbook workbook, StyleService styleService) {
        cell.setCellStyle(styleService.getDataCellStyle());

        if (value instanceof Date) {
            cell.setCellValue(DateUtils.getDateAsString((Date) value));
        } else if (value instanceof DateExcel) {
            DateExcel dateExcel = (DateExcel) value;
            cell.setCellValue(DateUtils.getDateAsString(dateExcel.getValue(), dateExcel.getFormat()));
            cell.setCellStyle(styleService.getCellStyle(workbook, DateExcel.class, dateExcel.getStyles()));
        } else if (value instanceof StringExcel) {
            StringExcel stringExcel = (StringExcel) value;
            cell.setCellValue(stringExcel.getValue());
            cell.setCellStyle(styleService.getCellStyle(workbook, StringExcel.class, stringExcel.getStyles()));
        } else if (value instanceof Number) {
            Number number = (Number) value;
            cell.setCellValue(number.getValue());
            cell.setCellStyle(styleService.getCellStyle(workbook, Number.class, number.getStyles()));
        } else if (value instanceof Boolean) {
            cell.setCellValue((Boolean) value);
        } else if (value instanceof Merge) {
            Merge merge = (Merge) value;
            String cellValue = merge.getValue();
            if (StringUtils.hasLength(cellValue)) {
                if (merge.getRange() > 1) {
                    if (merge.getOrientation() == Merge.Orientation.VERTICAL) {
                        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex + merge.getRange() - 1,
                                colIndex + merge.getOffset(), colIndex + merge.getOffset()));
                    } else if (merge.getOrientation() == Merge.Orientation.HORIZONTAL) {
                        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex,
                                colIndex + merge.getOffset(), colIndex + merge.getOffset() + merge.getRange() - 1));
                    }
                }
                cell = sheet.getRow(rowIndex).createCell(colIndex + merge.getOffset());
                cell.setCellValue(cellValue);
                cell.setCellStyle(styleService.getCellStyle(workbook, Merge.class, merge.getStyles()));
            }
        } else {
            cell.setCellValue(value.toString());
        }
    }

    // ── Read ──────────────────────────────────────────────────────────────────

    @Override
    public <T> List<T> readDynamicExcel(byte[] data, List<? extends ExcelHeaderBase> headers, Class<T> dataClass,
                                        ExcelReadSettings settings) throws ExcelGenerationException {
        try (Workbook workbook = WorkbookFactory.create(new ByteArrayInputStream(data))) {
            return readDynamicExcel(workbook, headers, dataClass, settings);
        } catch (ExcelGenerationException e) {
            throw e;
        } catch (IOException e) {
            log.error("Error opening excel from byte array!", e);
            throw new ExcelGenerationException("Error opening excel file.", e);
        }
    }

    @Override
    public <T> List<T> readDynamicExcel(Workbook workbook, List<? extends ExcelHeaderBase> headers, Class<T> dataClass,
                                        ExcelReadSettings settings) throws ExcelGenerationException {
        String sheetName = settings.getSheetName();
        if (sheetName == null) {
            throw new ExcelGenerationException("Sheet name must not be null in ExcelReadSettings.");
        }

        Sheet sheet = workbook.getSheet(sheetName);
        if (sheet == null) {
            throw new ExcelGenerationException("Sheet not found: " + sheetName);
        }

        try {
            dataClass.getDeclaredConstructor();
        } catch (NoSuchMethodException e) {
            throw new ExcelGenerationException(dataClass.getName() + " must have a no-args constructor.", e);
        }

        int headerRowIndex;
        List<String> searchKeys = settings.getSearchKeys();
        if (!searchKeys.isEmpty()) {
            headerRowIndex = findHeaderRow(sheet, searchKeys, settings.getColOffset(), settings.getMaxScanRows());
            if (headerRowIndex < 0) {
                throw new ExcelGenerationException("Could not find table header containing keys: " + searchKeys);
            }
            log.info("Found table header at row {} in sheet '{}'", headerRowIndex, sheetName);
        } else {
            headerRowIndex = settings.getRowOffset();
        }

        Row headerRow = sheet.getRow(headerRowIndex);
        if (headerRow == null) {
            throw new ExcelGenerationException("Header row is empty at row index: " + headerRowIndex);
        }

        Map<Integer, ExcelHeaderBase> columnMapping = buildColumnMapping(headerRow, headers, settings.getColOffset());
        if (columnMapping.isEmpty()) {
            log.warn("No columns matched the provided headers in sheet '{}'. Returning empty list.", sheetName);
            return new ArrayList<>();
        }

        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        List<T> result = new ArrayList<>();

        for (int rowIdx = headerRowIndex + 1; rowIdx <= sheet.getLastRowNum(); rowIdx++) {
            Row row = sheet.getRow(rowIdx);
            if (row == null || isRowEmpty(row)) continue;

            try {
                T instance = dataClass.getDeclaredConstructor().newInstance();
                boolean hasData = false;

                for (Map.Entry<Integer, ExcelHeaderBase> entry : columnMapping.entrySet()) {
                    Cell cell = row.getCell(entry.getKey());
                    Field field = dataClass.getDeclaredField(entry.getValue().getField());
                    field.setAccessible(true);
                    Object value = readCellValue(cell, field, evaluator);
                    if (value != null) {
                        field.set(instance, value);
                        hasData = true;
                    }
                }

                if (hasData) result.add(instance);
            } catch (ReflectiveOperationException e) {
                log.warn("Skipping row {}: could not map to {}", rowIdx, dataClass.getSimpleName(), e);
            }
        }

        log.info("Read {} rows from sheet '{}'", result.size(), sheetName);
        return result;
    }

    /**
     * Scans the sheet row by row to find the first row that contains all searchKeys as cell values.
     * Returns the row index, or -1 if not found within the scan limit.
     */
    private int findHeaderRow(Sheet sheet, List<String> searchKeys, int colOffset, int maxScanRows) {
        int limit = maxScanRows > 0 ? Math.min(maxScanRows, sheet.getLastRowNum() + 1) : sheet.getLastRowNum() + 1;
        Set<String> keys = new HashSet<>(searchKeys);

        for (int i = 0; i < limit; i++) {
            Row row = sheet.getRow(i);
            if (row == null) continue;

            Set<String> cellValues = new HashSet<>();
            for (Cell cell : row) {
                if (cell.getColumnIndex() >= colOffset && cell.getCellType() == CellType.STRING) {
                    cellValues.add(cell.getStringCellValue().trim());
                }
            }
            if (cellValues.containsAll(keys)) return i;
        }
        return -1;
    }

    /**
     * Builds a map of column index → ExcelHeaderBase by matching each header cell's string value
     * to the display names of the provided headers.
     */
    private Map<Integer, ExcelHeaderBase> buildColumnMapping(Row headerRow, List<? extends ExcelHeaderBase> headers,
                                                             int colOffset) {
        Map<String, ExcelHeaderBase> byDisplayName = new HashMap<>();
        for (ExcelHeaderBase header : headers) {
            if (header.getDisplayName() != null) {
                byDisplayName.put(header.getDisplayName(), header);
            }
        }

        Map<Integer, ExcelHeaderBase> columnMapping = new HashMap<>();
        for (Cell cell : headerRow) {
            if (cell.getColumnIndex() < colOffset || cell.getCellType() != CellType.STRING) continue;
            ExcelHeaderBase header = byDisplayName.get(cell.getStringCellValue().trim());
            if (header != null) {
                columnMapping.put(cell.getColumnIndex(), header);
            }
        }
        return columnMapping;
    }

    /**
     * Reads a cell's value and converts it to the type expected by the target field.
     * Supports String, Double, Integer, Long, Boolean, Date and the library's wrapper types.
     */
    private Object readCellValue(Cell cell, Field field, FormulaEvaluator evaluator) {
        if (cell == null) return null;

        CellType effectiveType = cell.getCellType();
        if (effectiveType == CellType.FORMULA) {
            effectiveType = evaluator.evaluateFormulaCell(cell);
        }

        Class<?> fieldType = field.getType();

        switch (effectiveType) {
            case STRING: {
                String value = cell.getStringCellValue().trim();
                if (value.isEmpty()) return null;
                if (fieldType == StringExcel.class) return StringExcel.fromValue(value);
                return value;
            }
            case NUMERIC: {
                if (DateUtil.isCellDateFormatted(cell)) {
                    Date value = cell.getDateCellValue();
                    if (fieldType == DateExcel.class) return DateExcel.fromValue(value);
                    return value;
                }
                double value = cell.getNumericCellValue();
                if (fieldType == Number.class) return Number.fromValue(value);
                if (fieldType == Integer.class || fieldType == int.class) return (int) value;
                if (fieldType == Long.class || fieldType == long.class) return (long) value;
                if (fieldType == String.class || fieldType == StringExcel.class) {
                    String str = value % 1 == 0 ? String.valueOf((long) value) : String.valueOf(value);
                    return fieldType == StringExcel.class ? StringExcel.fromValue(str) : str;
                }
                return value;
            }
            case BOOLEAN:
                return cell.getBooleanCellValue();
            case BLANK:
            default:
                return null;
        }
    }

    private boolean isRowEmpty(Row row) {
        for (Cell cell : row) {
            if (cell == null || cell.getCellType() == CellType.BLANK) continue;
            if (cell.getCellType() != CellType.STRING) return false;
            if (!cell.getStringCellValue().trim().isEmpty()) return false;
        }
        return true;
    }
}
