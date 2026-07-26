# Class \avadim\FastExcelLaravel\ExcelReader

<!-- Hand-maintained. The wrapper delegates via __call, which a reflection-based
     generator cannot represent, so do not auto-generate this page; edit it by hand
     when the class changes. -->

---

`ExcelReader` is a thin Laravel wrapper around a **FastExcelReader** book. It does not extend a concrete
reader; instead it **composes** an `avadim\FastExcelReader\AbstractBook` and adds the Eloquent-aware import
helpers. Because it relies only on the shared `AbstractBook` / `AbstractSheet` public API, the same wrapper
works for every format the reader supports — **XLSX, legacy XLS (Excel 97-2003) and CSV** — and the format is
detected automatically from the file content (see [Reading legacy XLS files](80-reading-xls.md) and
[Reading CSV files](81-reading-csv.md)).

Any method that is not listed below is transparently forwarded to the wrapped `AbstractBook` (see
[Delegated methods](#delegated-methods)).

> **Upgrade note.** Up to and including 3.x `ExcelReader` extended `\avadim\FastExcelReader\Excel`. It no
> longer does, so it is **not** an instance of that class anymore. Type-hint on
> `\avadim\FastExcelLaravel\ExcelReader` instead. Static helpers that used to be inherited
> (`validate()`, `colLetter()`, `colNum()`, `setTempDir()`, …) now live only on
> `\avadim\FastExcelReader\Excel`.

---

Own methods:

* [open()](#open) -- Open a workbook (XLSX or XLS) for import
* [sheet()](#sheet) -- Returns the current or a named sheet as a SheetReader
* [withHeadings()](#withheadings) -- Set headings for the current sheet
* [mapping()](#mapping) -- Set mapping callback for the current sheet
* [importModel()](#importmodel) -- Import data into a model from the current sheet
* [readRows()](#readrows) -- Read rows from the current sheet (applies the mapping, if any)
* [getBook()](#getbook) -- Returns the wrapped book instance
* [wrapSheet()](#wrapsheet) -- Wraps a book sheet in a SheetReader (used internally)
* [__construct()](#__construct)

---

## open()

---

```php
public static function open(string $file, ?array $options = []): ExcelReader
```
_Open a workbook for import. The format is selected from the file content, so XLSX, XLS and CSV are opened the
same way; the file extension is ignored. A file that is neither a valid XLSX (ZIP) nor XLS (OLE2) is read as CSV._

### Parameters

* `string $file`
* `array|null $options` -- supported keys:
  * `temp_dir` -- directory for temporary files used when reading XLSX; falls back to the
    `fast-excel.temp_dir` config value. XLS and CSV are read straight from the file and never use a temp dir.
  * CSV options -- `delimiter`, `enclosure`, `escape`, `encoding`, `skip_empty_lines`, `comment_prefix`,
    `mode` (`'strict'`/`'tolerant'`). They apply only to CSV and fall back to the `fast-excel.csv` config
    values; a per-call option wins, a `null` keeps the reader default. See
    [Reading CSV files](81-reading-csv.md).

---

## sheet()

---

```php
public function sheet(?string $name = null): ?SheetReader
```
_Returns the current sheet, or the sheet with the given name, wrapped in a [SheetReader](95-api-class-sheetreader.md).
Returns `null` if the named sheet does not exist._

### Parameters

* `string|null $name`

---

## withHeadings()

---

```php
public function withHeadings(?array $headers = []): ExcelReader
```
_Set headings for the current sheet. The first row of the read area is always skipped; if `$headers` is not
empty, these names are used as attribute names (in column order) instead of the values of the first row.
The setting applies to the next import operation only._

> **Naming.** This is the wrapper's name for the reader's own `withHeader()`
> (`avadim\FastExcelReader\AbstractSheet::withHeader()`). The wrapper uses `withHeadings()` to match the
> Laravel ecosystem convention (as in maatwebsite/excel); the underlying reader calls the same concept
> `withHeader()`.

### Parameters

* `array|null $headers`

---

## mapping()

---

```php
public function mapping($callback): ExcelReader
```
_Set a mapping for the current sheet. Accepts a callback `function (array $row): array` or an array of
`column => attribute` correspondences. The setting applies to `importModel()` and `readRows()`._

### Parameters

* `\Closure|callable|array $callback`

---

## importModel()

---

```php
public function importModel(string $modelClass, $address = null, $columns = null): ExcelReader
```
_Import data from the current sheet into a model: a new model is filled and saved for each row
(`fill()` + `save()`)._

### Parameters

* `string $modelClass`
* `string|bool|null $address` -- read area, e.g. `'B:D'`, `'B4'`, `'B4:D7'`
* `array|bool|null $columns`

---

## readRows()

---

```php
public function readRows($columnKeys = [], ?int $resultMode = null, ?bool $styleIdxInclude = null): array
```
_Returns cell values of the current sheet as a two-dimensional array `[row][col]`. If a mapping is set, it is
applied to every row._

### Parameters

* `array|bool|int|null $columnKeys`
* `int|null $resultMode`
* `bool|null $styleIdxInclude`

---

## getBook()

---

```php
public function getBook(): \avadim\FastExcelReader\AbstractBook
```
_Returns the wrapped book instance (an `avadim\FastExcelReader\Excel` for XLSX,
`avadim\FastExcelReader\Xls\XlsBook` for XLS or `avadim\FastExcelReader\Csv\CsvBook` for CSV)._

### Parameters

_None_

---

## wrapSheet()

---

```php
public function wrapSheet(\avadim\FastExcelReader\AbstractSheet $sheet): SheetReader
```
_Wraps a sheet returned by the book in a [SheetReader](95-api-class-sheetreader.md), reusing the wrapper so
that state set on a sheet (headings, mapping) survives later access to the same sheet. Called internally; you
normally use [sheet()](#sheet) instead._

### Parameters

* `\avadim\FastExcelReader\AbstractSheet $sheet`

---

## __construct()

---

```php
public function __construct(\avadim\FastExcelReader\AbstractBook $book)
```
_Wraps an already opened book. Use the static [open()](#open) factory instead of constructing directly._

### Parameters

* `\avadim\FastExcelReader\AbstractBook $book`

---

## Delegated methods

Every reading method not listed above is forwarded to the wrapped `AbstractBook` via `__call`. When a
delegated call returns a sheet (or a list of sheets), the result is re-wrapped in a
[SheetReader](95-api-class-sheetreader.md) so fluent chains keep working, e.g.:

```php
$excel->mapping($callback)->from('A2')->readRows();
```

The full list of delegated methods — `getSheetNames()`, `getSheet()`, `selectSheet()`, `setReadArea()`,
`from()`, `readCells()`, `readColumns()`, `readCallback()`, `getDefinedNames()`, `setDateFormat()`,
`dateFormatter()`, image and style helpers, and so on — is documented in the underlying **FastExcelReader**
library: https://github.com/aVadim483/fast-excel-reader#readme

---
