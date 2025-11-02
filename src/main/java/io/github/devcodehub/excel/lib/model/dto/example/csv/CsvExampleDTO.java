package io.github.devcodehub.excel.lib.model.dto.example.csv;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

/**
 * This class represents a row in the Csv Example
 */
@Getter
@Setter
@Builder
@AllArgsConstructor
@NoArgsConstructor
public class CsvExampleDTO implements CsvDTOBase {

    private String value1;
    private String value2;
    private Integer value3;
    private double value4;

    @Override
    public Object[] getFields() {
        return new Object[] {value1, value2, value3, value4};
    }
}
