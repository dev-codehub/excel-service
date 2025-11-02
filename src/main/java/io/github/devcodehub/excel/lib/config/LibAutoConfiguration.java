package io.github.devcodehub.excel.lib.config;

import io.github.devcodehub.excel.lib.service.ExcelService;
import io.github.devcodehub.excel.lib.service.ExcelServiceImpl;
import io.github.devcodehub.excel.lib.service.csv.CsvService;
import io.github.devcodehub.excel.lib.service.csv.CsvServiceImpl;
import org.springframework.boot.autoconfigure.AutoConfiguration;
import org.springframework.context.annotation.Bean;

@AutoConfiguration
public class LibAutoConfiguration {
    @Bean
    public ExcelService excelService() {
        return new ExcelServiceImpl();
    }

    @Bean
    public CsvService csvService() {
        return new CsvServiceImpl();
    }
}
