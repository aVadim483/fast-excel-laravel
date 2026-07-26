# Class \avadim\FastExcelLaravel\SheetReader

<!-- Hand-maintained. The wrapper delegates via __call, which a reflection-based
     generator cannot represent, so do not auto-generate this page; edit it by hand
     when the class changes. -->

---

`SheetReader` is a thin Laravel wrapper around a **FastExcelReader** sheet. It does not extend a concrete
sheet; instead it **composes** an `avadim\FastExcelReader\AbstractSheet` and adds the Eloquent-aware import
helpers. Because it relies only on the shared `AbstractSheet` public API, the same wrapper works for every
format the reader supports (XLSX, XLS, …).

Any method that is not listed below is transparently forwarded to the wrapped `AbstractSheet` (see
[Delegated methods](#delegated-methods)).

> **Upgrade note.** Up to and including 3.x `SheetReader` extended `\avadim\FastExcelReader\Sheet`. It no
> longer does, so it is **not** an instance of that class anymore. Type-hint on
> `\avadim\FastExcelLaravel\SheetReader` instead.

---

Own methods:

* [withHeadings()](#withheadings) -- Set headings for the sheet
* [mapping()](#mapping) -- Set mapping callback for the sheet
* [importModel()](#importmodel) -- Load models from the sheet into the database
* [readRows()](#readrows) -- Returns cell values as a two-dimensional array (applies the mapping, if any)
* [getSheet()](#getsheet) -- Returns the wrapped sheet instance
* [__construct()](#__construct)

---

## withHeadings()

---

```php
public function withHeadings(?array $headers = []): SheetReader
```
_Set headings for the sheet. The first row of the read area is always skipped; if `$headers` is not empty,
these names are used as attribute names (in column order) instead of the values of the first row. The setting
applies to the next import operation only._

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
public function mapping($callback): SheetReader
```
_Set a mapping for the sheet. Accepts a callback `function (array $row): array` or an array of
`column => attribute` correspondences. The setting applies to `importModel()` and `readRows()`._

### Parameters

* `\Closure|callable|array $callback`

---

## importModel()

---

```php
public function importModel($modelClass, $address = null, $columns = null): SheetReader
```
_Load models from the sheet into the database: a new model is filled and saved for each row
(`fill()` + `save()`)._

```php
importModel(User::class)                 // whole sheet
importModel(User::class, 'B:D')          // read columns B:D
importModel(User::class, 'B3')           // read area starting at B3
importModel(User::class, 'B3', true)     // read area starting at B3, first row as field names
```

### Parameters

* `$modelClass`
* `$address`
* `$columns`

---

## readRows()

---

```php
public function readRows($columnKeys = [], ?int $resultMode = null, ?bool $styleIdxInclude = null): array
```
_Returns cell values as a two-dimensional array `[row][col]`. If a mapping is set, it is applied to every row._

### Parameters

* `array|bool|int|null $columnKeys`
* `int|null $resultMode`
* `bool|null $styleIdxInclude`

---

## getSheet()

---

```php
public function getSheet(): \avadim\FastExcelReader\AbstractSheet
```
_Returns the wrapped sheet instance (an `avadim\FastExcelReader\Sheet` for XLSX or
`avadim\FastExcelReader\Xls\XlsSheet` for XLS)._

### Parameters

_None_

---

## __construct()

---

```php
public function __construct(\avadim\FastExcelReader\AbstractSheet $sheet, ExcelReader $excel)
```
_Wraps a sheet owned by an [ExcelReader](94-api-class-excelreader.md). Created internally by
`ExcelReader::sheet()` / `ExcelReader::wrapSheet()`; you normally do not construct it directly._

### Parameters

* `\avadim\FastExcelReader\AbstractSheet $sheet`
* `ExcelReader $excel`

---

## Delegated methods

Every reading method not listed above is forwarded to the wrapped `AbstractSheet` via `__call`, e.g.
`name()`, `nextRow()`, `setReadArea()`, `setReadAreaColumns()`, `from()`, `readCells()`, `readColumns()`,
`readFirstRow()`, `getMergedCells()`, `countRows()` / `countColumns()`, the dimension helpers, and the image
and style helpers. A sheet returned by a delegated call is re-wrapped so fluent chains keep working.

The full list is documented in the underlying **FastExcelReader** library:
https://github.com/aVadim483/fast-excel-reader#readme

---
