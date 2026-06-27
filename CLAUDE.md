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

# Install to local Maven repository
mvn install -DskipTests

# Run tests (none exist yet — to be added)
mvn test

# Deploy to Maven Central (requires GPG key + OSSRH credentials)
mvn deploy
```

There is no linter configured. Java 8 source/target compatibility is enforced by `maven-compiler-plugin`.

## Architecture

### Data flow

```
Caller
  └─► ExcelService (interface)
        └─► ExcelServiceImpl
              ├─► StyleService          (creates/caches POI CellStyle objects)
              └─► Apache POI Workbook   (XSSFWorkbook / .xlsx)
```

`ExcelServiceImpl.generateDynamicExcelWorkbook` is the core method. All other public methods delegate to it. The flow inside that method is:

1. Create/reuse a `Workbook` and `Sheet`.
2. Instantiate `StyleService` — it pre-builds default `CellStyle` objects for headers and each data type (`Merge`, `DateExcel`, `StringExcel`, `Number`).
3. Write **data rows first** (via reflection on `dataClass` fields), then write the **header row last** — this is intentional so that `autoSizeColumn` is driven by data width, not header width.
4. Apply column width constraints (min/max from `StyleDTO`).
5. Optionally apply auto-filter and freeze pane.

### Key contracts

**`ExcelHeaderBase`** — implemented as an enum by the caller. Each constant maps a POJO field name (`getField()`) to a display name (`getDisplayName()`) and optional per-column `StyleDTO`.

**`ExcelSettings`** — configures the sheet: name, row/column offsets, custom global styles (`ExcelCustomStyles`), filter, and freeze pane.

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
- Reflection (`getDeclaredField` + `setAccessible(true)`) is used to read private fields from data POJOs — field names in header enums must match exactly.
- The library targets Java 8; avoid using APIs introduced after Java 8.
