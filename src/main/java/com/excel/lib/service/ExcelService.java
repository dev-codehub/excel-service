package com.excel.lib.service;

import com.excel.lib.model.dto.excel.ExcelSettings;
import com.excel.lib.model.dto.excel.ExcelHeaderBase;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.List;

public interface ExcelService {
    public byte[] generateDynamicExcel(List<? extends ExcelHeaderBase> headers, List<?> data, Class<?> dataClass,
                                       ExcelSettings excelSettings, Workbook workbook) throws Exception;

    public Workbook generateDynamicExcelWorkbook(List<? extends ExcelHeaderBase> headers, List<?> data,
                                                 Class<?> dataClass, ExcelSettings excelSettings,
                                                 Workbook auxWorkbook) throws Exception;
}
