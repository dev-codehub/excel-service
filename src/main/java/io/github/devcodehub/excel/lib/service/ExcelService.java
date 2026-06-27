package io.github.devcodehub.excel.lib.service;

import io.github.devcodehub.excel.lib.model.dto.excel.ExcelHeaderBase;
import io.github.devcodehub.excel.lib.model.dto.excel.ExcelSettings;
import io.github.devcodehub.excel.lib.model.dto.exception.ExcelGenerationException;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.List;

public interface ExcelService {
    byte[] generateDynamicExcel(List<? extends ExcelHeaderBase> headers, List<?> data, Class<?> dataClass,
                                ExcelSettings excelSettings) throws ExcelGenerationException;

    byte[] generateDynamicExcel(List<? extends ExcelHeaderBase> headers, List<?> data, Class<?> dataClass,
                                ExcelSettings excelSettings, Workbook workbook) throws ExcelGenerationException;

    Workbook generateDynamicExcelWorkbook(List<? extends ExcelHeaderBase> headers, List<?> data,
                                          Class<?> dataClass, ExcelSettings excelSettings) throws ExcelGenerationException;

    Workbook generateDynamicExcelWorkbook(List<? extends ExcelHeaderBase> headers, List<?> data,
                                          Class<?> dataClass, ExcelSettings excelSettings,
                                          Workbook workbook) throws ExcelGenerationException;
}
