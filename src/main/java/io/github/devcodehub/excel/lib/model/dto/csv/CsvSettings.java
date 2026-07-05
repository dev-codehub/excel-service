package io.github.devcodehub.excel.lib.model.dto.csv;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;

@Getter
@Setter
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class CsvSettings {
    @Builder.Default
    private char delimiter = ';';
    @Builder.Default
    private Charset charset = StandardCharsets.UTF_8;
}
