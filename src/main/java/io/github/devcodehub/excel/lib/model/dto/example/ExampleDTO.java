package io.github.devcodehub.excel.lib.model.dto.example;

import io.github.devcodehub.excel.lib.model.dto.excel.datatype.Merge;
import io.github.devcodehub.excel.lib.model.dto.excel.datatype.Number;
import io.github.devcodehub.excel.lib.model.dto.excel.datatype.StringExcel;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

/**
 * This class represents a row in the Example Excel
 */
@Getter
@Setter
@Builder
@AllArgsConstructor
@NoArgsConstructor
public class ExampleDTO {
    private StringExcel value1;
    private String value2;
    private Number value3;
    private Merge value4;
}
