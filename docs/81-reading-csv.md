# Reading CSV files

Since the reader dependency `avadim/fast-excel-reader` `^4.3`, this package reads **CSV files** in addition
to XLSX and legacy XLS.

## How it works

A CSV file has no signature. `\Excel::open()` first checks for the XLSX (ZIP) and XLS (OLE2) signatures; a
file that matches neither is read as **CSV**:

```php
// A .csv (or any non-XLSX/non-XLS file) is read as CSV
$excel = \Excel::open(storage_path('users.csv'));

// The whole import API is identical regardless of the source format
$excel->withHeadings()->importModel(User::class);
$rows = $excel->readRows();
```

So there is **nothing CSV-specific to call** — headings, mapping, `importModel()`, read areas, `readRows()`,
`nextRow()` and the other reading methods behave the same for CSV, XLS and XLSX. See
[Class ExcelReader](94-api-class-excelreader.md) and [Class SheetReader](95-api-class-sheetreader.md).

A CSV workbook is presented as a **single sheet** named `CSV`.

## CSV options

Options are passed as the second argument of `\Excel::open()` and forwarded to the reader:

```php
$excel = \Excel::open($file, [
    'delimiter'        => ';',        // column delimiter; null = auto-detect
    'enclosure'        => '"',        // field enclosure character
    'escape'           => '',         // escape character; '' = none
    'encoding'         => 'CP1251',   // source encoding; null = auto-detect (BOM is handled automatically)
    'skip_empty_lines' => true,       // skip blank lines
    'comment_prefix'   => '#',        // skip lines starting with this prefix
    'mode'             => 'tolerant', // 'strict' or 'tolerant' (ragged rows)
]);
```

Global defaults can be set once in `config/fast-excel.php`:

```php
'csv' => [
    'delimiter' => ';',
    'encoding'  => 'CP1251',
],
```

A per-call option always wins over a config default; a `null` value is skipped so the reader keeps its own
default. A UTF-8 BOM is detected and stripped automatically, so it never leaks into the first column.

## What is supported

* All values are read as **strings** — a CSV file carries no cell types, so there is no automatic
  number/date typing (cast in your model or in a `mapping()` callback if needed).
* `withHeadings()`, custom headings, `mapping()` (array or callback), read areas, `readRows()`,
  `nextRow()`, `importModel()` — all work exactly as for XLSX/XLS.
* `tolerant` mode reads ragged rows: short rows get `null` for the missing columns, longer rows keep the
  extra ones.

## Limitations

* **Writing CSV is not supported.** This package writes XLSX only (`\Excel::create()` always produces an
  XLSX workbook). CSV is a read-only format here.
* **No styles, formats, formulas, images or merged cells** — a CSV file has none of them; the corresponding
  reading methods return empty results rather than throwing.
* **Detection is by elimination.** Because CSV has no signature, any file that is not a valid XLSX (ZIP) or
  XLS (OLE2) is treated as delimited text — including a mislabeled `.xlsx` that is actually plain text.

For the full CSV feature matrix and option list see the underlying library:
https://github.com/aVadim483/fast-excel-reader#readme
