package io.github.devcodehub.excel.lib.service.csv;

import io.github.devcodehub.excel.lib.model.dto.example.csv.CsvDTOBase;
import io.github.devcodehub.excel.lib.model.dto.example.csv.CsvHeaderBase;
import java.util.List;

public interface CsvService {

    <T extends Enum<T> & CsvHeaderBase> byte[] generateCsv(Class<T> headerEnumClass,
                                                                  final List<? extends CsvDTOBase> data,
                                                                  final Class<? extends CsvDTOBase> dataClass)
            throws Exception;
}
