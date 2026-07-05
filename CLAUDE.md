# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

`excel-service` is a Java 8 library published on Maven Central that generates Excel workbooks dynamically using Apache POI. It is designed as a Spring Boot auto-configurable library — consumers inject `ExcelService` as a Spring bean and call a single method to produce an `.xlsx` file.

## Build Commands

```bash
# Compile
mvn compile

# Package (produces JAR in target/)
mvn package -DskipTests

# Install to local Maven repository (works without GPG — signing is release-only)
mvn clean install

# Run all tests (JUnit 5)
mvn test

# Run a single test class or method
mvn test -Dtest=ExcelServiceImplReadTest
mvn test -Dtest=ExcelServiceImplReadTest#readDynamicExcel_sheetNotFound_throwsExcelGenerationException

# Deploy + sign for Maven Central (requires GPG + Central credentials)
mvn clean deploy -Prelease
```

There is no linter configured. Java 8 source/target compatibility is enforced by `maven-compiler-plugin`.

GPG signing lives in a `release` Maven profile, so a plain `mvn clean install` runs on machines without GPG installed. Signing only happens when `-Prelease` is passed (used for Maven Central publishing).

## Architecture

The library is bidirectional: it **writes** POJO lists to `.xlsx` and **reads** `.xlsx` back into POJO lists. Both directions live in `ExcelServiceImpl` and are driven by the same `ExcelHeaderBase` enum (field ⇄ display-name mapping).

### Write data flow

```
Caller
  └─► ExcelService (interface)
        └─► ExcelServiceImpl
              ├─► StyleService          (creates/caches POI CellStyle objects)
              └─► Apache POI Workbook   (XSSFWorkbook / .xlsx)
```

`ExcelServiceImpl.generateDynamicExcelWorkbook` is the core write method. All other write methods delegate to it. The flow inside that method is:

1. Create/reuse a `Workbook` and `Sheet`.
2. Instantiate `StyleService` — it pre-builds default `CellStyle` objects for headers and each data type (`Merge`, `DateExcel`, `StringExcel`, `Number`).
3. Write **data rows first** (via reflection on `dataClass` fields), then write the **header row last** — this is intentional so that `autoSizeColumn` is driven by data width, not header width.
4. Apply column width constraints (min/max from `StyleDTO`).
5. Optionally apply auto-filter and freeze pane.

### Read data flow

`ExcelServiceImpl.readDynamicExcel(Workbook, …)` is the core read method; the `byte[]` overload opens a workbook via `WorkbookFactory` and delegates to it. The flow is:

1. Locate the header row: if `ExcelReadSettings.searchKeys` is non-empty, scan rows (up to `maxScanRows`) for the first row containing all keys; otherwise use `rowOffset` directly.
2. Build a `column index → ExcelHeaderBase` map by matching header cell text to each header's **display name** (so column order in the file is irrelevant).
3. For each data row, instantiate `dataClass` (requires a **no-args constructor**), then set each mapped field via reflection. `readCellValue` coerces the POI cell type to the target field type, resolving formulas via a `FormulaEvaluator` and unwrapping into `StringExcel`/`DateExcel`/`Number` when the field is a wrapper type (using their `fromValue` factories).
4. Rows that are empty or fail to map are skipped (logged), not fatal.

### Key contracts

**`ExcelHeaderBase`** — implemented as an enum by the caller. Each constant maps a POJO field name (`getField()`) to a display name (`getDisplayName()`) and optional per-column `StyleDTO`.

**`ExcelSettings`** — configures a **write**: sheet name, row/column offsets, custom global styles (`ExcelCustomStyles`), filter, and freeze pane.

**`ExcelReadSettings`** — configures a **read**: sheet name, `rowOffset`/`colOffset`, `searchKeys` (locate the table by content instead of a fixed offset), and `maxScanRows`. Sheet name is required; a missing sheet or unlocatable header throws `ExcelGenerationException`.

**Data type wrappers** (`DateExcel`, `StringExcel`, `Number`, `Merge`) — wrap raw values and carry an optional `StyleDTO` for per-cell styling. If a field's value is a plain Java type (`String`, `Boolean`, `Date`), it is handled directly without a wrapper.

**`StyleService`** — instantiated fresh per `generateDynamicExcelWorkbook` call because POI `CellStyle` objects are bound to a specific `Workbook` instance. It holds one default style per data type and creates new styles on demand via `getNewCellStyle` (clone → override).

### Style resolution order (most specific wins)

```
Per-cell StyleDTO (on wrapper)  >  Per-column StyleDTO (on header enum)  >  ExcelCustomStyles (global defaults)
```

### `Merge` type behaviour

`Merge` allows merging multiple cells either vertically or horizontally. The `offset` field positions the merged cell relative to the current column index, and `range` defines how many cells to span.

### Spring auto-configuration

`LibAutoConfiguration` registers `ExcelServiceImpl` as a `@Lazy` bean. The auto-configuration is declared in `META-INF/spring/org.springframework.boot.autoconfigure.AutoConfiguration.imports` for Spring Boot 2.7+ compatibility.

## Conventions

- All model classes use Lombok (`@Getter`, `@Setter`, `@Builder`, `@NoArgsConstructor`, `@AllArgsConstructor`).
- Colors are represented via `ExcelColor` (enum) or any `ExcelColorBase` implementation, backed by XSSF RGB byte arrays.
- POI width unit: 1 character = 256 units. Constants `DEFAULT_COLUMN_WIDTH`, `MAX_COLUMN_WIDTH`, and `DEFAULT_COLUMN_MARGIN` in `ExcelServiceImpl` control auto-sizing bounds.
- Reflection (`getDeclaredField` + `setAccessible(true)`) is used to read/write private fields on data POJOs — field names in header enums must match exactly. Read additionally requires a no-args constructor on `dataClass`.
- All public methods throw the checked `ExcelGenerationException`; user-facing messages come from the `ResponseCode` enum.
- The library targets Java 8; avoid using APIs introduced after Java 8.

## Tests

JUnit 5 (`junit-jupiter`), run by `maven-surefire-plugin`. Tests live under `src/test/java` mirroring the main package layout, split by concern: `ExcelServiceImplWriteTest`, `ExcelServiceImplReadTest`, `StyleServiceTest`, plus utility tests (`DateUtilsTest`, `ExcelUtilsTest`). Write/read tests round-trip through real in-memory POI workbooks (no mocking of POI). `src/main/.../model/dto/example/` (`ExampleDTO`, `ExcelExampleHeader`) provides reusable fixtures for both tests and documentation.
