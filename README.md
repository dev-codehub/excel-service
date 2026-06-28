# Excel Service

A Java library that generates and reads Excel workbooks dynamically using Apache POI — with a single method call.

[![Maven Central](https://img.shields.io/maven-central/v/io.github.dev-codehub/excel-service)](https://central.sonatype.com/artifact/io.github.dev-codehub/excel-service)
[![License](https://img.shields.io/badge/license-Apache%202.0-blue.svg)](http://www.apache.org/licenses/LICENSE-2.0)

---

## Table of Contents

1. [Requirements](#requirements)
2. [Technologies](#technologies)
3. [Setup](#setup)
4. [Features](#features)
5. [Quick Start](#quick-start)
6. [Writing Excel Files](#writing-excel-files)
   - [Data Types](#data-types)
   - [ExcelSettings](#excelsettings)
   - [Styling](#styling)
   - [Merging Cells](#merging-cells)
7. [Reading Excel Files](#reading-excel-files)
   - [ExcelReadSettings](#excelreadsettings)
   - [Locating the Table by Keys](#locating-the-table-by-keys)
8. [API Reference](#api-reference)

---

## Requirements

- Java 8 or later
- Spring Boot 2.x or later (auto-configuration is optional)

## Technologies

- Java 8
- Spring Boot 2 (auto-configuration)
- Apache POI 5.3.0

---

## Setup

Add the dependency to your `pom.xml`:

```xml
<dependency>
    <groupId>io.github.dev-codehub</groupId>
    <artifactId>excel-service</artifactId>
    <version>1.0.3</version>
</dependency>
```

`ExcelService` is auto-configured as a Spring bean — inject it directly:

```java
@Service
public class ReportService {
    private final ExcelService excelService;

    public ReportService(ExcelService excelService) {
        this.excelService = excelService;
    }
}
```

---

## Features

| Feature | Description |
|---|---|
| Write | Generate `.xlsx` files from a list of DTOs |
| Read | Parse `.xlsx` files back into a list of DTOs |
| Styling | Per-cell, per-column, and global styles |
| Merging | Horizontal and vertical cell merging |
| Auto-filter | Enable column filter dropdowns on the header row |
| Freeze pane | Lock rows/columns while scrolling |
| Multiple sheets | Append new sheets to an existing workbook |
| Table detection | Locate data tables by searching for header key words |

---

## Quick Start

### Define your DTO

```java
@Getter @Setter @Builder @NoArgsConstructor @AllArgsConstructor
public class PersonDTO {
    private String name;
    private String email;
    private Number salary;   // use Number wrapper for numeric cells
}
```

### Define your header enum

```java
@AllArgsConstructor
public enum PersonHeader implements ExcelHeaderBase {
    NAME("name",   "Name",   null),
    EMAIL("email", "Email",  null),
    SALARY("salary", "Salary", null);

    private final String field;
    private final String displayName;
    private final StyleDTO styles;

    @Override public String getField()       { return field; }
    @Override public String getDisplayName() { return displayName; }
    @Override public StyleDTO getStyles()    { return styles; }
}
```

### Write

```java
List<PersonDTO> people = repository.findAll();

ExcelSettings settings = ExcelSettings.builder()
        .sheetName("People")
        .headerFilterActive(true)
        .build();

byte[] bytes = excelService.generateDynamicExcel(
        Arrays.asList(PersonHeader.values()), people, PersonDTO.class, settings);
```

### Read

```java
ExcelReadSettings readSettings = ExcelReadSettings.builder()
        .sheetName("People")
        .build();

List<PersonDTO> people = excelService.readDynamicExcel(
        bytes, Arrays.asList(PersonHeader.values()), PersonDTO.class, readSettings);
```

---

## Writing Excel Files

### Data Types

| Java type | Excel cell | Notes |
|---|---|---|
| `String` | String | Written as-is |
| `Boolean` | Boolean | Native boolean cell |
| `Date` | String | Formatted as `yyyy/MM/dd` |
| `StringExcel` | String | Same as `String` but supports per-cell styling |
| `Number` | Numeric | Use for proper numeric cells (supports formulas, sorting) |
| `DateExcel` | String | `Date` with a custom format pattern |
| `Merge` | String | Merges a range of cells horizontally or vertically |
| Any other type | String | `toString()` is called |

**`StringExcel`** — string cell with optional style:
```java
StringExcel.fromValue("Hello")
StringExcel.fromValue("Hello", StyleDTO.builder().bold(true).build())
```

**`Number`** — numeric cell with optional style:
```java
Number.fromValue(1234.56)
Number.fromValue(1234.56, StyleDTO.builder().foregroundColor(ExcelColor.LIGHT_GREEN).build())
```

**`DateExcel`** — date cell with a custom format:
```java
DateExcel.fromValue(new Date(), "dd/MM/yyyy")
DateExcel.fromValue(new Date(), "dd/MM/yyyy", myStyleDTO)
```

---

### ExcelSettings

```java
ExcelSettings settings = ExcelSettings.builder()
        .sheetName("Report")          // required
        .rowOffset(2)                 // first data row starts at rowOffset + 1 (default: 0)
        .colOffset(1)                 // all columns shifted right by colOffset (default: 0)
        .headerFilterActive(true)     // enable auto-filter on the header row
        .freezePane(ExcelSettings.FreezePane.builder()
                .colSplit(1)          // freeze first column
                .rowSplit(1)          // freeze first row
                .build())
        .excelCustomStyles(ExcelCustomStyles.builder()
                .headerBold(true)
                .headerBorderActive(true)
                .headerBorderColor(ExcelColor.BLUE)
                .fontFamily("Arial")
                .build())
        .build();
```

**`ExcelSettings` fields:**

| Field | Type | Default | Description |
|---|---|---|---|
| `sheetName` | `String` | — | Sheet name (required) |
| `rowOffset` | `int` | `0` | Row index where the header row is written |
| `colOffset` | `int` | `0` | Column index where the first column starts |
| `headerFilterActive` | `boolean` | `false` | Adds a dropdown filter to the header row |
| `freezePane` | `FreezePane` | `null` | Freezes rows/columns |
| `excelCustomStyles` | `ExcelCustomStyles` | defaults | Global style overrides |

---

### Styling

Styles are resolved from most specific to least specific:

```
Per-cell StyleDTO (on wrapper)  >  Per-column StyleDTO (on header enum)  >  ExcelCustomStyles (global)
```

**`StyleDTO` fields:**

| Field | Type | Description |
|---|---|---|
| `bold` | `Boolean` | Bold font |
| `fontHeight` | `Double` | Font size in points |
| `foregroundColor` | `ExcelColorBase` | Cell background colour |
| `textColor` | `ExcelColorBase` | Font colour |
| `horizontalAlignment` | `HorizontalAlignment` | POI enum (LEFT, CENTER, RIGHT…) |
| `verticalAlignment` | `VerticalAlignment` | POI enum (TOP, CENTER, BOTTOM…) |
| `minWidth` | `Integer` | Minimum column width in characters |
| `maxWidth` | `Integer` | Maximum column width in characters |

**`ExcelCustomStyles` fields (global defaults):**

| Field | Type | Default | Description |
|---|---|---|---|
| `headerFontHeight` | `short` | `11` | Header font size |
| `headersHeight` | `Double` | `24.0` | Header row height in points |
| `headerBold` | `boolean` | `false` | Bold header font |
| `headerBorderActive` | `boolean` | `true` | Show border on header cells |
| `headerBorderColor` | `ExcelColorBase` | `BLACK` | Header border colour |
| `dataCellFontHeight` | `short` | `11` | Data cell font size |
| `dataCellBold` | `boolean` | `false` | Bold data font |
| `dataBorderActive` | `boolean` | `false` | Show border on data cells |
| `dataBorderColor` | `ExcelColorBase` | `LIGHT_GREY` | Data cell border colour |
| `fontFamily` | `String` | `"Calibri"` | Font family for all cells |

**Available colours** (`ExcelColor` enum):

`BLACK` · `BLUE` · `BLUE_ACCENT_1_LIGHTER_40` · `LIGHT_BLUE` · `BROWN` · `LIGHT_BROWN` · `LIGHT_CYAN` · `GREEN` · `LIGHT_GREEN` · `DARK_GREY` · `GREY` · `LIGHT_GREY` · `DARK_MINT` · `LIGHT_DARK_MINT` · `ORANGE` · `LIGHT_ORANGE` · `PURPLE` · `LIGHT_PURPLE` · `PINK` · `LIGHT_PINK` · `RED` · `LIGHT_RED` · `TEAL` · `LIGHT_TEAL` · `WHITE` · `YELLOW` · `LIGHT_YELLOW` · `YELLOW_BROWN` · `LIGHT_YELLOW_BROWN`

You can also implement `ExcelColorBase` to supply any custom RGB colour.

---

### Merging Cells

**Horizontal merge** — spans columns to the right:

```java
// Merge 3 cells horizontally, starting at the current column
Merge.fromValue(3, 0, "Merged Header", Merge.Orientation.HORIZONTAL)
```

**Vertical merge** — spans rows downward:

```java
// Merge 2 rows vertically
Merge.fromValue(2, 0, "Group Label", Merge.Orientation.VERTICAL)
```

`Merge.fromValue(int range, int offset, String value, Orientation orientation)`:
- `range` — number of cells to span (must be > 1 to create a merged region)
- `offset` — column offset relative to the current column index
- `value` — text written in the first cell of the merged region

---

### Multiple Sheets

Pass an existing `Workbook` to append a new sheet without losing previous content:

```java
Workbook workbook = excelService.generateDynamicExcelWorkbook(
        headers1, data1, DTO1.class, settings1);

workbook = excelService.generateDynamicExcelWorkbook(
        headers2, data2, DTO2.class, settings2, workbook);

// Serialize to bytes when done
ByteArrayOutputStream out = new ByteArrayOutputStream();
workbook.write(out);
workbook.close();
byte[] bytes = out.toByteArray();
```

---

## Reading Excel Files

Parse an `.xlsx` file back into a list of DTOs using the same header enum used to write it.

```java
ExcelReadSettings readSettings = ExcelReadSettings.builder()
        .sheetName("People")
        .build();

List<PersonDTO> people = excelService.readDynamicExcel(
        bytes, Arrays.asList(PersonHeader.values()), PersonDTO.class, readSettings);
```

The reader maps columns by matching the **display name** of each header constant to the text found in the header row of the sheet. Only columns present in both the enum and the sheet are populated; unrecognised columns are ignored.

**Cell type mapping on read:**

| Excel cell type | Java field type | Result |
|---|---|---|
| String | `String` | String value |
| String | `StringExcel` | `StringExcel.fromValue(...)` |
| Numeric (integer) | `Integer` / `Long` | Cast from double |
| Numeric (decimal) | `Double` | Raw double |
| Numeric (decimal) | `Number` | `Number.fromValue(...)` |
| Boolean | `Boolean` | Boolean value |
| Formula | any | Evaluated before mapping |
| Blank / null | any | Field left as `null` |

> **Note:** Plain `Double` fields written by the library are stored as string cells (no `Number` wrapper) and cannot be read back into a `Double` field. Use the `Number` wrapper type for numeric fields that need to survive a write/read round-trip.

---

### ExcelReadSettings

| Field | Type | Default | Description |
|---|---|---|---|
| `sheetName` | `String` | — | Sheet to read (required) |
| `rowOffset` | `int` | `0` | Row index of the header row when `searchKeys` is empty |
| `colOffset` | `int` | `0` | Ignore columns to the left of this index |
| `searchKeys` | `List<String>` | `[]` | Display names used to locate the header row automatically |
| `maxScanRows` | `int` | `0` | Scan limit when using `searchKeys` (0 = no limit) |

---

### Locating the Table by Keys

When a sheet contains preamble content (titles, metadata) before the actual data table, use `searchKeys` to let the reader find the header row automatically:

```java
ExcelReadSettings readSettings = ExcelReadSettings.builder()
        .sheetName("Report")
        .searchKeys(Arrays.asList("Name", "Email", "Salary"))
        .maxScanRows(20)   // stop scanning after 20 rows
        .build();

List<PersonDTO> people = excelService.readDynamicExcel(
        bytes, Arrays.asList(PersonHeader.values()), PersonDTO.class, readSettings);
```

The reader scans each row until it finds one whose string cells contain **all** the specified keys. That row becomes the header row, and data is read from the next row onward. An `ExcelGenerationException` is thrown if the keys are not found within the scan limit.

---

## API Reference

```java
// Write — returns byte array
byte[] generateDynamicExcel(headers, data, dataClass, settings) throws ExcelGenerationException;

// Write — appends to an existing workbook
byte[] generateDynamicExcel(headers, data, dataClass, settings, workbook) throws ExcelGenerationException;

// Write — returns Workbook for multi-sheet scenarios
Workbook generateDynamicExcelWorkbook(headers, data, dataClass, settings) throws ExcelGenerationException;
Workbook generateDynamicExcelWorkbook(headers, data, dataClass, settings, workbook) throws ExcelGenerationException;

// Read — from byte array
<T> List<T> readDynamicExcel(byte[] data, headers, dataClass, settings) throws ExcelGenerationException;

// Read — from an open Workbook (caller manages lifecycle)
<T> List<T> readDynamicExcel(Workbook workbook, headers, dataClass, settings) throws ExcelGenerationException;
```

All methods throw `ExcelGenerationException` (a checked exception) on validation errors, missing sheets, or I/O failures.
