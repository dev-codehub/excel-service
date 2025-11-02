package io.github.devcodehub.excel.lib.service.csv;

import io.github.devcodehub.excel.lib.model.dto.example.csv.CsvDTOBase;
import io.github.devcodehub.excel.lib.model.dto.example.csv.CsvHeaderBase;
import io.github.devcodehub.excel.lib.model.dto.exception.ResponseCode;
import io.github.devcodehub.excel.lib.utils.CsvHeaderUtils;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.List;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.csv.CSVFormat.Builder;
import org.apache.commons.csv.CSVPrinter;
import org.springframework.context.annotation.Lazy;
import org.springframework.stereotype.Service;
import org.springframework.util.CollectionUtils;

@Lazy
@Slf4j
@Service
public class CsvServiceImpl implements CsvService {

    @Override
    public <T extends Enum<T> & CsvHeaderBase> byte[] generateCsv(Class<T> headerEnumClass,
                                                                  final List<? extends CsvDTOBase> data,
                                                                  final Class<? extends CsvDTOBase> dataClass)
            throws Exception {
        if (CollectionUtils.isEmpty(data)) {
            throw new IllegalArgumentException("DTO list cannot be empty");
        }

        try {
            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
            OutputStreamWriter writer = new OutputStreamWriter(outputStream);
            CSVPrinter csvPrinter = new CSVPrinter(writer,
                    Builder.create()
                           .setHeader(CsvHeaderUtils.getHeaders(headerEnumClass))
                           .setDelimiter(';')
                           .build()
            );

            for (CsvDTOBase dto : data) {
                csvPrinter.printRecord(dto.getFields());
            }

            csvPrinter.flush();
            return outputStream.toByteArray();
        } catch (IOException e) {
            log.error("Error generating CSV file!");
            throw new Exception(ResponseCode.ERROR_GENERATING_CSV.getMessage());
        }
    }
}
