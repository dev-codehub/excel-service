package io.github.devcodehub.excel.lib.service;

import io.github.devcodehub.excel.lib.model.dto.excel.ExcelSettings;
import io.github.devcodehub.excel.lib.model.dto.excel.ExcelHeaderBase;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.List;

public interface ExcelService {
    public byte[] generateDynamicExcel(List<? extends ExcelHeaderBase> headers, List<?> data, Class<?> dataClass,
                                       ExcelSettings excelSettings) throws Exception;
    public byte[] generateDynamicExcel(List<? extends ExcelHeaderBase> headers, List<?> data, Class<?> dataClass,
                                       ExcelSettings excelSettings, Workbook workbook) throws Exception;

    public Workbook generateDynamicExcelWorkbook(List<? extends ExcelHeaderBase> headers, List<?> data,
                                                 Class<?> dataClass, ExcelSettings excelSettings) throws Exception;

    public Workbook generateDynamicExcelWorkbook(List<? extends ExcelHeaderBase> headers, List<?> data,
                                                 Class<?> dataClass, ExcelSettings excelSettings,
                                                 Workbook workbook) throws Exception;
}
