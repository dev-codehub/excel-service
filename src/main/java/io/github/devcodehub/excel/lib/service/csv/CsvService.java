package io.github.devcodehub.excel.lib.service.csv;

import io.github.devcodehub.excel.lib.model.dto.csv.CsvHeaderBase;
import io.github.devcodehub.excel.lib.model.dto.csv.CsvSettings;
import io.github.devcodehub.excel.lib.model.dto.exception.CsvGenerationException;

import java.util.List;

public interface CsvService {

    byte[] generateCsv(List<? extends CsvHeaderBase> headers,
                       List<?> data,
                       Class<?> dataClass,
                       CsvSettings settings) throws CsvGenerationException;
}
