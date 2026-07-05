package io.github.devcodehub.excel.lib.service.csv;

import io.github.devcodehub.excel.lib.model.dto.csv.CsvHeaderBase;
import io.github.devcodehub.excel.lib.model.dto.csv.CsvSettings;
import io.github.devcodehub.excel.lib.model.dto.exception.CsvGenerationException;
import io.github.devcodehub.excel.lib.model.dto.exception.ResponseCode;
import io.github.devcodehub.excel.lib.utils.CsvHeaderUtils;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVPrinter;
import org.apache.commons.csv.CSVRecord;
import org.springframework.util.CollectionUtils;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.OutputStreamWriter;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

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

    @Override
    public <T> List<T> readCsv(byte[] csv,
                                List<? extends CsvHeaderBase> headers,
                                Class<T> dataClass,
                                CsvSettings settings) throws CsvGenerationException {
        final java.lang.reflect.Constructor<T> noArgsCtor;
        try {
            noArgsCtor = dataClass.getDeclaredConstructor();
            noArgsCtor.setAccessible(true);
        } catch (NoSuchMethodException e) {
            throw new CsvGenerationException(dataClass.getName() + " must have a no-args constructor.", e);
        }

        Map<String, CsvHeaderBase> byDisplayName = new HashMap<>();
        for (CsvHeaderBase header : headers) {
            byDisplayName.put(header.getDisplayName(), header);
        }

        try (InputStreamReader reader = new InputStreamReader(new ByteArrayInputStream(csv), settings.getCharset());
             CSVParser csvParser = new CSVParser(reader,
                     CSVFormat.Builder.create()
                             .setDelimiter(settings.getDelimiter())
                             .build())) {

            List<CSVRecord> records = csvParser.getRecords();
            if (records.isEmpty()) {
                return new ArrayList<>();
            }

            // First record is the header row — build column index → header mapping
            CSVRecord headerRecord = records.get(0);
            Map<Integer, CsvHeaderBase> columnMapping = new HashMap<>();
            for (int i = 0; i < headerRecord.size(); i++) {
                CsvHeaderBase header = byDisplayName.get(headerRecord.get(i).trim());
                if (header != null) {
                    columnMapping.put(i, header);
                }
            }

            if (columnMapping.isEmpty()) {
                log.warn("No columns matched the provided headers. Returning empty list.");
                return new ArrayList<>();
            }

            List<T> result = new ArrayList<>();

            for (int r = 1; r < records.size(); r++) {
                CSVRecord record = records.get(r);
                try {
                    T instance = noArgsCtor.newInstance();
                    boolean hasData = false;

                    for (Map.Entry<Integer, CsvHeaderBase> entry : columnMapping.entrySet()) {
                        if (entry.getKey() >= record.size()) continue;
                        String rawValue = record.get(entry.getKey());
                        if (rawValue == null || rawValue.isEmpty()) continue;

                        Field field = dataClass.getDeclaredField(entry.getValue().getField());
                        field.setAccessible(true);
                        Object value = convertValue(rawValue, field.getType());
                        if (value != null) {
                            field.set(instance, value);
                            hasData = true;
                        }
                    }

                    if (hasData) result.add(instance);
                } catch (ReflectiveOperationException e) {
                    log.warn("Skipping record {}: could not map to {}", r + 1, dataClass.getSimpleName(), e);
                }
            }

            log.info("Read {} records from CSV", result.size());
            return result;
        } catch (IOException e) {
            log.error("Error reading CSV file", e);
            throw new CsvGenerationException(ResponseCode.ERROR_READING_CSV.getMessage(), e);
        }
    }

    private Object convertValue(String value, Class<?> targetType) {
        try {
            if (targetType == String.class) return value;
            if (targetType == Integer.class || targetType == int.class) return Integer.parseInt(value.trim());
            if (targetType == Long.class || targetType == long.class) return Long.parseLong(value.trim());
            if (targetType == Double.class || targetType == double.class) return Double.parseDouble(value.trim());
            if (targetType == Float.class || targetType == float.class) return Float.parseFloat(value.trim());
            if (targetType == Boolean.class || targetType == boolean.class) return Boolean.parseBoolean(value.trim());
            return value;
        } catch (NumberFormatException e) {
            log.warn("Could not convert '{}' to {}", value, targetType.getSimpleName());
            return null;
        }
    }
}
