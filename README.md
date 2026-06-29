# Excel Service

> Gera e lê ficheiros Excel a partir de listas de DTOs Java com uma única chamada de método — sem boilerplate, sem configuração manual do Apache POI.

[![Maven Central](https://img.shields.io/maven-central/v/io.github.dev-codehub/excel-service?color=1a7f37)](https://central.sonatype.com/artifact/io.github.dev-codehub/excel-service)
[![Java](https://img.shields.io/badge/Java-8%2B-blue)](#)
[![Apache POI](https://img.shields.io/badge/Apache%20POI-5.3.0-blue)](#)
[![License](https://img.shields.io/badge/License-Apache%202.0-lightgrey)](http://www.apache.org/licenses/LICENSE-2.0)

---

## Índice

- [Setup](#setup)
- [Quick Start](#quick-start)
- [Escrita de ficheiros Excel](#escrita-de-ficheiros-excel)
  - [Tipos de dados](#tipos-de-dados)
  - [ExcelSettings](#excelsettings)
  - [Sistema de estilos](#sistema-de-estilos)
  - [Merge de células](#merge-de-células)
  - [Múltiplas folhas](#múltiplas-folhas)
- [Leitura de ficheiros Excel](#leitura-de-ficheiros-excel)
  - [ExcelReadSettings](#excelreadsettings)
  - [Deteção automática de tabelas](#deteção-automática-de-tabelas)
  - [Mapeamento de tipos](#mapeamento-de-tipos-na-leitura)
- [API Reference](#api-reference)

---

## Setup

```xml
<dependency>
    <groupId>io.github.dev-codehub</groupId>
    <artifactId>excel-service</artifactId>
    <version>1.0.3</version>
</dependency>
```

A biblioteca regista-se automaticamente como bean Spring Boot — injeta diretamente:

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

## Quick Start

### 1. Define o DTO

```java
@Getter @Setter @Builder @NoArgsConstructor @AllArgsConstructor
public class PersonDTO {
    private String  name;
    private String  email;
    private Number  salary;    // wrapper Number → célula numérica
    private Date    birthDate; // Date → formatado como yyyy/MM/dd
    private Boolean active;    // Boolean → célula booleana nativa
}
```

### 2. Define o enum de cabeçalhos

O enum implementa `ExcelHeaderBase` e mapeia cada campo do DTO ao cabeçalho visível na coluna. O `field` tem de corresponder exatamente ao nome da propriedade no DTO.

```java
@AllArgsConstructor
public enum PersonHeader implements ExcelHeaderBase {
    NAME       ("name",      "Nome",            null),
    EMAIL      ("email",     "Email",           null),
    SALARY     ("salary",    "Salário",         null),
    BIRTH_DATE ("birthDate", "Data Nascimento", null),
    ACTIVE     ("active",    "Ativo",           null);

    private final String   field;
    private final String   displayName;
    private final StyleDTO styles;

    @Override public String   getField()       { return field; }
    @Override public String   getDisplayName() { return displayName; }
    @Override public StyleDTO getStyles()      { return styles; }
}
```

### 3. Escreve e lê

```java
List<ExcelHeaderBase> headers = Arrays.asList(PersonHeader.values());

// Escrever
byte[] bytes = excelService.generateDynamicExcel(
        headers, personList, PersonDTO.class,
        ExcelSettings.builder().sheetName("Pessoas").headerFilterActive(true).build());

// Ler
List<PersonDTO> pessoas = excelService.readDynamicExcel(
        bytes, headers, PersonDTO.class,
        ExcelReadSettings.builder().sheetName("Pessoas").build());
```

---

## Escrita de ficheiros Excel

### Tipos de dados

O tipo Java da propriedade no DTO determina o tipo de célula criada. Para controlo total sobre o tipo de célula ou o estilo, usa os **wrappers** da biblioteca.

| Tipo Java | Tipo de célula Excel | Notas |
|---|---|---|
| `String` | Texto | Escrito diretamente |
| `Boolean` | Booleano | Célula booleana nativa |
| `Date` | Texto | Formatado como `yyyy/MM/dd` |
| `StringExcel` | Texto | Igual a `String` mas com estilo por célula |
| `Number` | Numérico | Usar para somas, fórmulas e ordenação numérica |
| `DateExcel` | Texto | `Date` com formato personalizado |
| `Merge` | Texto | Funde células horizontal ou verticalmente |
| Outro qualquer | Texto | `toString()` é chamado |

> **Atenção:** Campos `Double` nativos são escritos como células de texto. Para células numéricas com round-trip de leitura, usa o wrapper `Number`.

**Exemplos de criação dos wrappers:**

```java
// StringExcel — texto com estilo opcional por célula
StringExcel.fromValue("Aprovado")
StringExcel.fromValue("Aprovado", StyleDTO.builder().foregroundColor(ExcelColor.LIGHT_GREEN).build())

// Number — célula numérica
Number.fromValue(1_500.75)
Number.fromValue(1_500.75, StyleDTO.builder().horizontalAlignment(HorizontalAlignment.RIGHT).build())

// DateExcel — data com formato personalizado
DateExcel.fromValue(new Date())                       // usa yyyy/MM/dd
DateExcel.fromValue(new Date(), "dd/MM/yyyy")         // formato personalizado
DateExcel.fromValue(new Date(), "dd/MM/yyyy", style)  // com estilo
```

---

### ExcelSettings

```java
ExcelSettings settings = ExcelSettings.builder()
        .sheetName("Relatório")           // obrigatório
        .rowOffset(2)                     // cabeçalho na linha 2, dados a partir da 3
        .colOffset(1)                     // todas as colunas avançam 1 para a direita
        .headerFilterActive(true)         // dropdown de filtro no cabeçalho
        .freezePane(ExcelSettings.FreezePane.builder()
                .colSplit(1)              // congela a primeira coluna
                .rowSplit(1)             // congela a primeira linha
                .build())
        .excelCustomStyles(ExcelCustomStyles.builder()
                .headerBold(true)
                .fontFamily("Arial")
                .build())
        .build();
```

| Campo | Tipo | Default | Descrição |
|---|---|---|---|
| `sheetName` | `String` | — | Nome da folha **(obrigatório)** |
| `rowOffset` | `int` | `0` | Índice da linha onde o cabeçalho é escrito |
| `colOffset` | `int` | `0` | Índice da coluna onde a primeira coluna começa |
| `headerFilterActive` | `boolean` | `false` | Dropdown de filtro na linha de cabeçalho |
| `freezePane` | `FreezePane` | `null` | Congela linhas e/ou colunas durante o scroll |
| `excelCustomStyles` | `ExcelCustomStyles` | *defaults* | Estilos globais da folha |

---

### Sistema de estilos

Os estilos são resolvidos por especificidade — o mais específico prevalece:

```
StyleDTO na célula  >  StyleDTO no cabeçalho (enum)  >  ExcelCustomStyles (global)
```

**`StyleDTO`** — controlo por célula ou por coluna:

| Campo | Tipo | Descrição |
|---|---|---|
| `bold` | `Boolean` | Texto a negrito |
| `fontHeight` | `Double` | Tamanho da fonte em pontos |
| `foregroundColor` | `ExcelColorBase` | Cor de fundo da célula |
| `textColor` | `ExcelColorBase` | Cor do texto |
| `horizontalAlignment` | `HorizontalAlignment` | LEFT, CENTER, RIGHT… |
| `verticalAlignment` | `VerticalAlignment` | TOP, CENTER, BOTTOM… |
| `minWidth` | `Integer` | Largura mínima da coluna em caracteres |
| `maxWidth` | `Integer` | Largura máxima da coluna em caracteres |

**`ExcelCustomStyles`** — defaults globais definidos em `ExcelSettings`:

| Campo | Default | Descrição |
|---|---|---|
| `headerFontHeight` | `11` | Tamanho da fonte do cabeçalho |
| `headersHeight` | `24.0` | Altura da linha de cabeçalho em pontos |
| `headerBold` | `false` | Cabeçalho a negrito |
| `headerBorderActive` | `true` | Borda nas células de cabeçalho |
| `headerBorderColor` | `BLACK` | Cor da borda do cabeçalho |
| `dataCellFontHeight` | `11` | Tamanho da fonte dos dados |
| `dataCellBold` | `false` | Dados a negrito |
| `dataBorderActive` | `false` | Borda nas células de dados |
| `dataBorderColor` | `LIGHT_GREY` | Cor da borda dos dados |
| `fontFamily` | `"Calibri"` | Família de fontes para todas as células |

**`ExcelColor`** — cores disponíveis:

<details>
<summary>Ver paleta completa</summary>

| Constante | Constante | Constante |
|---|---|---|
| `BLACK` | `BLUE` | `BLUE_ACCENT_1_LIGHTER_40` |
| `LIGHT_BLUE` | `BROWN` | `LIGHT_BROWN` |
| `LIGHT_CYAN` | `GREEN` | `LIGHT_GREEN` |
| `DARK_GREY` | `GREY` | `LIGHT_GREY` |
| `DARK_MINT` | `LIGHT_DARK_MINT` | `ORANGE` |
| `LIGHT_ORANGE` | `PURPLE` | `LIGHT_PURPLE` |
| `PINK` | `LIGHT_PINK` | `RED` |
| `LIGHT_RED` | `TEAL` | `LIGHT_TEAL` |
| `WHITE` | `YELLOW` | `LIGHT_YELLOW` |
| `YELLOW_BROWN` | `LIGHT_YELLOW_BROWN` | |

Implementa `ExcelColorBase` para cores RGB personalizadas.

</details>

---

### Merge de células

```java
// Horizontal — funde 3 colunas a partir da posição atual
Merge.fromValue(3, 0, "Dados Pessoais", Merge.Orientation.HORIZONTAL)

// Vertical — funde 2 linhas a partir da linha atual
Merge.fromValue(2, 0, "Grupo A", Merge.Orientation.VERTICAL)

// Com offset — começa 1 coluna à frente da posição atual do campo
Merge.fromValue(2, 1, "Sub-cabeçalho", Merge.Orientation.HORIZONTAL)
```

| Parâmetro | Descrição |
|---|---|
| `range` | Número de células a fundir (> 1 para criar região fundida) |
| `offset` | Deslocamento de coluna relativo à posição atual do campo no DTO |
| `value` | Texto escrito na primeira célula da região fundida |
| `orientation` | `HORIZONTAL` (expande colunas) ou `VERTICAL` (expande linhas) |

---

### Múltiplas folhas

```java
// Primeira folha — cria um Workbook novo
Workbook wb = excelService.generateDynamicExcelWorkbook(
        headers1, data1, Dto1.class,
        ExcelSettings.builder().sheetName("Vendas").build());

// Segunda folha — reutiliza o mesmo Workbook
wb = excelService.generateDynamicExcelWorkbook(
        headers2, data2, Dto2.class,
        ExcelSettings.builder().sheetName("Clientes").build(), wb);

// Serializar para bytes
ByteArrayOutputStream out = new ByteArrayOutputStream();
wb.write(out);
wb.close();
byte[] bytes = out.toByteArray();
```

---

## Leitura de ficheiros Excel

```java
List<ExcelHeaderBase> headers = Arrays.asList(PersonHeader.values());

ExcelReadSettings settings = ExcelReadSettings.builder()
        .sheetName("Pessoas")
        .build();

List<PersonDTO> pessoas = excelService.readDynamicExcel(
        bytes, headers, PersonDTO.class, settings);
```

A biblioteca mapeia colunas comparando o texto das células do cabeçalho com o `displayName` de cada enum. Colunas desconhecidas são ignoradas. O DTO deve ter um construtor sem argumentos.

---

### ExcelReadSettings

| Campo | Tipo | Default | Descrição |
|---|---|---|---|
| `sheetName` | `String` | — | Folha a ler **(obrigatório)** |
| `rowOffset` | `int` | `0` | Linha de cabeçalho (usado quando `searchKeys` está vazio) |
| `colOffset` | `int` | `0` | Ignora colunas à esquerda deste índice |
| `searchKeys` | `List<String>` | `[]` | Nomes de colunas para localizar o cabeçalho automaticamente |
| `maxScanRows` | `int` | `0` | Limite de linhas a percorrer com `searchKeys` (0 = sem limite) |

---

### Deteção automática de tabelas

Quando a folha tem conteúdo antes da tabela (títulos, metadados), usa `searchKeys` para localizar o cabeçalho sem saber o índice da linha.

```java
// A folha pode ter preamble antes da tabela:
// Linha 0: "Relatório Anual"
// Linha 1: "Gerado em: 2024-01-15"
// Linha 2: (vazia)
// Linha 3: "Nome" | "Email" | "Salário"  ← cabeçalho encontrado aqui
// Linha 4+: dados

ExcelReadSettings settings = ExcelReadSettings.builder()
        .sheetName("Pessoas")
        .searchKeys(Arrays.asList("Nome", "Email"))  // basta um subconjunto único
        .maxScanRows(20)                              // para após 20 linhas
        .build();
```

> Se as `searchKeys` não forem encontradas dentro do limite `maxScanRows`, é lançado `ExcelGenerationException`.

---

### Mapeamento de tipos na leitura

| Célula Excel | Tipo do campo no DTO | Resultado |
|---|---|---|
| Texto | `String` | Valor como string |
| Texto | `StringExcel` | `StringExcel.fromValue(...)` |
| Numérico (inteiro) | `Integer` / `int` | Cast de double para int |
| Numérico (inteiro) | `Long` / `long` | Cast de double para long |
| Numérico (decimal) | `Double` / `double` | Valor double bruto |
| Numérico | `Number` | `Number.fromValue(...)` |
| Numérico (data) | `Date` | Data nativa |
| Numérico (data) | `DateExcel` | `DateExcel.fromValue(...)` |
| Booleano | `Boolean` | Valor booleano |
| Fórmula | qualquer | Avaliada antes do mapeamento |
| Em branco / null | qualquer | Campo fica `null` |

---

## API Reference

Todos os métodos lançam `ExcelGenerationException` (checked) em caso de validação inválida, folha não encontrada ou erro de I/O.

### Escrita

```java
// Gera e devolve bytes prontos a enviar ou guardar
byte[] generateDynamicExcel(
    List<? extends ExcelHeaderBase> headers,
    List<?> data,
    Class<?> dataClass,
    ExcelSettings settings) throws ExcelGenerationException;

// Igual, mas adiciona uma nova folha a um Workbook existente
byte[] generateDynamicExcel(
    List<? extends ExcelHeaderBase> headers,
    List<?> data,
    Class<?> dataClass,
    ExcelSettings settings,
    Workbook workbook) throws ExcelGenerationException;

// Devolve o Workbook POI — útil para cenários de múltiplas folhas
Workbook generateDynamicExcelWorkbook(
    List<? extends ExcelHeaderBase> headers,
    List<?> data,
    Class<?> dataClass,
    ExcelSettings settings) throws ExcelGenerationException;

// Adiciona uma folha a um Workbook existente e devolve-o
Workbook generateDynamicExcelWorkbook(
    List<? extends ExcelHeaderBase> headers,
    List<?> data,
    Class<?> dataClass,
    ExcelSettings settings,
    Workbook workbook) throws ExcelGenerationException;
```

### Leitura

```java
// Analisa bytes de um .xlsx e devolve lista de DTOs
<T> List<T> readDynamicExcel(
    byte[] data,
    List<? extends ExcelHeaderBase> headers,
    Class<T> dataClass,
    ExcelReadSettings settings) throws ExcelGenerationException;

// Analisa um Workbook já aberto (ciclo de vida é responsabilidade do chamador)
<T> List<T> readDynamicExcel(
    Workbook workbook,
    List<? extends ExcelHeaderBase> headers,
    Class<T> dataClass,
    ExcelReadSettings settings) throws ExcelGenerationException;
```

---

## Licença

Distribuído sob a [Apache License 2.0](http://www.apache.org/licenses/LICENSE-2.0).
