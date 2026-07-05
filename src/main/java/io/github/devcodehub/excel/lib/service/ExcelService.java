package io.github.devcodehub.excel.lib.service;

import io.github.devcodehub.excel.lib.model.dto.excel.ExcelHeaderBase;
import io.github.devcodehub.excel.lib.model.dto.excel.ExcelReadSettings;
import io.github.devcodehub.excel.lib.model.dto.excel.ExcelSettings;
import io.github.devcodehub.excel.lib.model.dto.exception.ExcelGenerationException;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.List;

public interface ExcelService {

    // ── Write ────────────────────────────────────────────────────────────────

    /**
     * Generates an {@code .xlsx} workbook from a list of data objects and returns it as a byte array.
     *
     * @param headers      the column definitions mapping POJO field names to display names and styles
     * @param data         the rows to write; each element must be an instance of {@code dataClass}
     * @param dataClass    the class of the data objects, used to read field values via reflection
     * @param excelSettings the sheet configuration (name, offsets, styles, filter, freeze pane)
     * @return the generated workbook serialized as a byte array
     * @throws ExcelGenerationException if the workbook cannot be built or serialized
     */
    byte[] generateDynamicExcel(List<? extends ExcelHeaderBase> headers, List<?> data, Class<?> dataClass,
                                ExcelSettings excelSettings) throws ExcelGenerationException;

    /**
     * Generates a sheet from a list of data objects into an existing workbook and returns the result
     * as a byte array. Use this to write multiple sheets into the same file.
     *
     * @param headers      the column definitions mapping POJO field names to display names and styles
     * @param data         the rows to write; each element must be an instance of {@code dataClass}
     * @param dataClass    the class of the data objects, used to read field values via reflection
     * @param excelSettings the sheet configuration (name, offsets, styles, filter, freeze pane)
     * @param workbook     the existing workbook to add the sheet to
     * @return the resulting workbook serialized as a byte array
     * @throws ExcelGenerationException if the workbook cannot be built or serialized
     */
    byte[] generateDynamicExcel(List<? extends ExcelHeaderBase> headers, List<?> data, Class<?> dataClass,
                                ExcelSettings excelSettings, Workbook workbook) throws ExcelGenerationException;

    /**
     * Generates an {@code .xlsx} workbook from a list of data objects and returns the live
     * {@link Workbook} instance, leaving serialization and lifecycle to the caller.
     *
     * @param headers      the column definitions mapping POJO field names to display names and styles
     * @param data         the rows to write; each element must be an instance of {@code dataClass}
     * @param dataClass    the class of the data objects, used to read field values via reflection
     * @param excelSettings the sheet configuration (name, offsets, styles, filter, freeze pane)
     * @return the generated {@link Workbook}
     * @throws ExcelGenerationException if the workbook cannot be built
     */
    Workbook generateDynamicExcelWorkbook(List<? extends ExcelHeaderBase> headers, List<?> data,
                                          Class<?> dataClass, ExcelSettings excelSettings) throws ExcelGenerationException;

    /**
     * Adds a sheet built from a list of data objects to an existing workbook and returns the same
     * {@link Workbook} instance. This is the core method all other write methods delegate to.
     *
     * @param headers      the column definitions mapping POJO field names to display names and styles
     * @param data         the rows to write; each element must be an instance of {@code dataClass}
     * @param dataClass    the class of the data objects, used to read field values via reflection
     * @param excelSettings the sheet configuration (name, offsets, styles, filter, freeze pane)
     * @param workbook     the existing workbook to add the sheet to
     * @return the same {@link Workbook} with the new sheet added
     * @throws ExcelGenerationException if the workbook cannot be built
     */
    Workbook generateDynamicExcelWorkbook(List<? extends ExcelHeaderBase> headers, List<?> data,
                                          Class<?> dataClass, ExcelSettings excelSettings,
                                          Workbook workbook) throws ExcelGenerationException;

    // ── Read ─────────────────────────────────────────────────────────────────

    /**
     * Parses an Excel file from a byte array into a list of DTOs.
     * The table is located either by a fixed row offset or by scanning for header keys,
     * as configured in {@link ExcelReadSettings}.
     *
     * @param <T>       the type of the DTO produced for each row
     * @param data      the raw {@code .xlsx} file contents
     * @param headers   the column definitions mapping display names back to POJO field names
     * @param dataClass the class to instantiate and populate for each row
     * @param settings  the read configuration (sheet, header location strategy, offsets)
     * @return the parsed rows as a list of {@code dataClass} instances
     * @throws ExcelGenerationException if the file cannot be read or mapped to {@code dataClass}
     */
    <T> List<T> readDynamicExcel(byte[] data, List<? extends ExcelHeaderBase> headers, Class<T> dataClass,
                                 ExcelReadSettings settings) throws ExcelGenerationException;

    /**
     * Parses an Excel sheet from an open {@link Workbook} into a list of DTOs.
     * Caller is responsible for the workbook lifecycle.
     *
     * @param <T>       the type of the DTO produced for each row
     * @param workbook  the open workbook to read from
     * @param headers   the column definitions mapping display names back to POJO field names
     * @param dataClass the class to instantiate and populate for each row
     * @param settings  the read configuration (sheet, header location strategy, offsets)
     * @return the parsed rows as a list of {@code dataClass} instances
     * @throws ExcelGenerationException if the sheet cannot be read or mapped to {@code dataClass}
     */
    <T> List<T> readDynamicExcel(Workbook workbook, List<? extends ExcelHeaderBase> headers, Class<T> dataClass,
                                 ExcelReadSettings settings) throws ExcelGenerationException;
}
