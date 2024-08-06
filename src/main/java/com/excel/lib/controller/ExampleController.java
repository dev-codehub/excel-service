package com.excel.lib.controller;

import com.excel.lib.model.dto.example.ExampleDTO;
import com.excel.lib.model.dto.example.ExcelExampleHeader;
import com.excel.lib.model.dto.excel.ExcelColor;
import com.excel.lib.model.dto.excel.ExcelSettings;
import com.excel.lib.model.dto.excel.StyleDTO;
import com.excel.lib.model.dto.excel.datatype.Merge;
import com.excel.lib.model.dto.excel.datatype.Number;
import com.excel.lib.model.dto.excel.datatype.StringExcel;
import com.excel.lib.service.ExcelService;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;

import java.util.List;

@Controller
public class ExampleController {
    private final ExcelService excelService;

    public ExampleController(ExcelService excelService) {
        this.excelService = excelService;
    }

    @GetMapping("/download/example")
    public ResponseEntity<Resource> downloadExampleExcel() throws Exception {
        // Headers
        List<ExcelExampleHeader> headers = List.of(ExcelExampleHeader.values());

        // List of objects
        StyleDTO name1Style = StyleDTO.builder()
                .foregroundColor(ExcelColor.LIGHT_GREY)
                .bold(true)
                .build();
        Merge mergeCell = Merge.fromValue(4, 0, "Merge cell", Merge.Orientation.HORIZONTAL);
        Merge mergeCellVertical = Merge.fromValue(2, 0, "Merge cell", Merge.Orientation.VERTICAL);
        ExampleDTO example1 = ExampleDTO.builder()
                .value1(StringExcel.fromValue("Value1", name1Style))
                .value2("value2")
                .value3(Number.fromValue(15000.0, StyleDTO.builder().format("# ### ##0.00;[Red]-# ### ##0.00").build()))
                .value4(mergeCell)
                .build();
        ExampleDTO example2 = ExampleDTO.builder().value1(StringExcel.fromValue("Value1", name1Style)).value2("value2").value3(Number.fromValue(300000.0, StyleDTO.builder().format("# ### ##0.00;[Red]-# ### ##0.00").build())).value4(mergeCell).build();
        ExampleDTO example3 = ExampleDTO.builder().value1(StringExcel.fromValue("Value1", name1Style)).value2("value2").value3(Number.fromValue(-4500000.0, StyleDTO.builder().format("# ### ##0.00;[Red]-# ### ##0.00").build())).value4(mergeCellVertical).build();
        List<ExampleDTO> exampleDTOList = List.of(example1, example2, example3);

        // Excel settings
        ExcelSettings excelSettings = ExcelSettings.builder()
                .sheetName("example")
                .rowOffset(1)
                .colOffset(1)
                .freezePane(ExcelSettings.FreezePane.builder().rowSplit(2).colSplit(2).build())
                .build();

        // Generate excel
        byte[] bytes = excelService.generateDynamicExcel(headers, exampleDTOList, ExampleDTO.class, excelSettings, null);
        ByteArrayResource resource = new ByteArrayResource(bytes);

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=example.xlsx")
                .contentType(MediaType.APPLICATION_OCTET_STREAM)
                .body(resource);
    }
}
