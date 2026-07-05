package io.github.devcodehub.excel.lib.service.csv;

import io.github.devcodehub.excel.lib.model.dto.csv.CsvHeaderBase;
import io.github.devcodehub.excel.lib.model.dto.csv.CsvSettings;
import io.github.devcodehub.excel.lib.model.dto.exception.CsvGenerationException;
import io.github.devcodehub.excel.lib.model.dto.exception.ResponseCode;
import io.github.devcodehub.excel.lib.utils.CsvHeaderUtils;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;
import org.springframework.util.CollectionUtils;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

@Slf4j
public class CsvServiceImpl implements CsvService {

    @Override
    public byte[] generateCsv(List<? extends CsvHeaderBase> headers,
                               List<?> data,
                               Class<?> dataClass,
                               CsvSettings settings) throws CsvGenerationException {
        if (CollectionUtils.isEmpty(data)) {
            throw new CsvGenerationException(ResponseCode.ERROR_GENERATING_CSV.getMessage());
        }

        try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
             OutputStreamWriter writer = new OutputStreamWriter(outputStream, settings.getCharset());
             CSVPrinter csvPrinter = new CSVPrinter(writer,
                     CSVFormat.Builder.create()
                             .setHeader(CsvHeaderUtils.getDisplayNames(headers))
                             .setDelimiter(settings.getDelimiter())
                             .build())) {

            for (Object dto : data) {
                List<Object> values = new ArrayList<>();
                for (CsvHeaderBase header : headers) {
                    Field field = dataClass.getDeclaredField(header.getField());
                    field.setAccessible(true);
                    values.add(field.get(dto));
                }
                csvPrinter.printRecord(values);
            }

            csvPrinter.flush();
            return outputStream.toByteArray();
        } catch (IOException | NoSuchFieldException | IllegalAccessException e) {
            log.error("Error generating CSV file", e);
            throw new CsvGenerationException(ResponseCode.ERROR_GENERATING_CSV.getMessage(), e);
        }
    }
}
