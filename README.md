# Excel Service

A Java library to create Excel workbooks dynamically, with a single line of code you can create an Excel workbook.

Now you ask: "**That's impossible, I spend so much time and need huge amounts of code to generate an Excel!**"

Correction, you needed it, from now on it won't happen again.

# Table of Contents
1. [Requirements](#requirements)
1. [Technologies](#technologies)
1. [Features](#features)
1. [Setup](#setup)
1. [Examples](#examples)

# Requirements
- Requires Java 8 or later.

# Technologies
- Java 8
- Spring Boot 2
- Apache POI

# Features
- Generate excel workbook dynamically
- ~~Read dinamically excel files~~

# Setup
1. Add dependency to your project
   ```xml
        <!-- Excel Service -->
        <dependency>
            <groupId>io.github.dev-codehub</groupId>
            <artifactId>excel-service</artifactId>
            <version>1.0.3</version>
        </dependency>
   ```
2. Add Excel Service to your services
   ```java
    class ExampleService {
        private final ExcelService excelService;
   
        public ExampleService(ExcelService excelService) {
            this.excelService = excelService;
        }
    }
   ```

# Examples:
1. Simple Excel file with just one table and without custom styles
```java
    class ExampleService {
        private final ExcelService excelService;
   
        public ExampleService(ExcelService excelService) {
            this.excelService = excelService;
        }
   
        public Resource generateExampleExcel() {
            // Headers
            List<ExcelHeaderBase> headers = List.of(ExcelExampleHeader.values());

            // List of objects
            List<ExampleDTO> exampleDTOList = repository.getExampleDTOList();

            // Excel settings
            ExcelSettings excelSettings = ExcelSettings.builder()
                .sheetName("example")
                // Multiple settings and styles available
                .build();

            // Generate excel
            byte[] bytes = excelService.generateDynamicExcel(headers, exampleDTOList, ExampleDTO.class, excelSettings);

            //...
        }
    }
    
    @Getter
    @Setter
    @Builder
    @AllArgsConstructor
    @NoArgsConstructor
    public class ExampleDTO {
        private String value1;
        private String value2;
        private Number value3;
        private Merge value4;
    }
   
    @AllArgsConstructor
    public enum ExcelExampleHeader implements ExcelHeaderBase {
        VALUE1("value1", "Display Column 1"),
        VALUE2("value2", "Display Column 2"),
        VALUE3("value3", "Display Column 3"),
        VALUE4("value4", "Display Column 4");
    
        private final String field;
        private final String displayName;
        
        //...
    }
   ```

