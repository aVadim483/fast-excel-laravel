# Reading legacy XLS files

Since the reader dependency `avadim/fast-excel-reader` `^4.0`, this package reads the **legacy binary XLS
format (Excel 97-2003, BIFF8)** in addition to XLSX.

## How it works

You do not choose the format — `\Excel::open()` detects it automatically from the file **signature**, not from
the extension:

```php
// Both work exactly the same way; a file named .xlsx that is really an XLS is handled correctly too
$excel = \Excel::open(storage_path('report.xls'));
$excel = \Excel::open(storage_path('report.xlsx'));

// The whole import API is identical regardless of the source format
$excel->withHeadings()->importModel(User::class);
$rows = $excel->readRows();
```

So there is **nothing XLS-specific to call** — headings, mapping, `importModel()`, read areas, `readRows()`,
`nextRow()` and the other reading methods behave the same for XLS and XLSX. See
[Class ExcelReader](94-api-class-excelreader.md) and [Class SheetReader](95-api-class-sheetreader.md).

## What is supported

When reading XLS, the underlying reader extracts:

* cell values and types (dates are detected through the number format);
* cell styles — fonts, fills, borders, alignment, number formats and palette colours;
* formula text (the `f` field carries the leading `=`, exactly as for XLSX);
* embedded images;
* multiple worksheets, including hidden and very hidden ones, merged cells and sheet dimensions.

## Limitations

* **Writing XLS is not supported.** This package writes XLSX only (`\Excel::create()` always produces an XLSX
  workbook). XLS is a read-only format here.
* **Only Excel 97-2003 (BIFF8).** Older BIFF5/BIFF7 workbooks (Excel 5.0/95) and encrypted workbooks are
  rejected with an explanatory exception.
* Reading XLS never uses a temporary directory — the `temp_dir` option (and the `fast-excel.temp_dir` config
  value) only affects XLSX reading.

For the full XLS feature matrix see the underlying library:
https://github.com/aVadim483/fast-excel-reader#readme
