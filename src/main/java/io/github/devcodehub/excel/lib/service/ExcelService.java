package io.github.devcodehub.excel.lib.service;

import io.github.devcodehub.excel.lib.model.dto.excel.ExcelHeaderBase;
import io.github.devcodehub.excel.lib.model.dto.excel.ExcelReadSettings;
import io.github.devcodehub.excel.lib.model.dto.excel.ExcelSettings;
import io.github.devcodehub.excel.lib.model.dto.exception.ExcelGenerationException;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.List;

public interface ExcelService {

    // ── Write ────────────────────────────────────────────────────────────────

    byte[] generateDynamicExcel(List<? extends ExcelHeaderBase> headers, List<?> data, Class<?> dataClass,
                                ExcelSettings excelSettings) throws ExcelGenerationException;

    byte[] generateDynamicExcel(List<? extends ExcelHeaderBase> headers, List<?> data, Class<?> dataClass,
                                ExcelSettings excelSettings, Workbook workbook) throws ExcelGenerationException;

    Workbook generateDynamicExcelWorkbook(List<? extends ExcelHeaderBase> headers, List<?> data,
                                          Class<?> dataClass, ExcelSettings excelSettings) throws ExcelGenerationException;

    Workbook generateDynamicExcelWorkbook(List<? extends ExcelHeaderBase> headers, List<?> data,
                                          Class<?> dataClass, ExcelSettings excelSettings,
                                          Workbook workbook) throws ExcelGenerationException;

    // ── Read ─────────────────────────────────────────────────────────────────

    /**
     * Parses an Excel file from a byte array into a list of DTOs.
     * The table is located either by a fixed row offset or by scanning for header keys,
     * as configured in {@link ExcelReadSettings}.
     */
    <T> List<T> readDynamicExcel(byte[] data, List<? extends ExcelHeaderBase> headers, Class<T> dataClass,
                                 ExcelReadSettings settings) throws ExcelGenerationException;

    /**
     * Parses an Excel sheet from an open {@link Workbook} into a list of DTOs.
     * Caller is responsible for the workbook lifecycle.
     */
    <T> List<T> readDynamicExcel(Workbook workbook, List<? extends ExcelHeaderBase> headers, Class<T> dataClass,
                                 ExcelReadSettings settings) throws ExcelGenerationException;
}
