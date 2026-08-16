# Reading from a string or a stream

A workbook does not have to be a file on the local disk. Besides `\Excel::open()` the package provides two
more entry points for import:

* `\Excel::openString($content)` — the workbook is held in a **string** (a database blob, an HTTP response
  body, `Storage::get()`, a mail attachment, …);
* `\Excel::openStream($stream)` — the workbook is behind an open **stream resource** (`Storage::readStream()`,
  `fopen('https://…')`, `php://memory`, an uploaded file handle, …).

## How it works

Both methods write the content into a temporary file and then open it exactly as `\Excel::open()` does, so:

* the **format is detected from the content** — XLSX, legacy XLS and CSV all work, and no file name or
  extension is involved at all (see [Reading legacy XLS files](80-reading-xls.md) and
  [Reading CSV files](81-reading-csv.md));
* the **second argument is the same** as for `open()` — `temp_dir` and the CSV options
  (`delimiter`, `enclosure`, `escape`, `encoding`, `skip_empty_lines`, `comment_prefix`, `mode`), with the
  same fallback to `config('fast-excel.temp_dir')` and `config('fast-excel.csv')`;
* the **whole import API** is available afterwards — `withHeadings()`, `mapping()`, `importModel()`, read
  areas, `readRows()`, `nextRow()`, and everything delegated to the underlying book.

```php
// A workbook stored in a database column
$excel = \Excel::openString($report->file_body);
$excel->withHeadings()->importModel(User::class);

// A CSV from a Laravel disk, with options
$stream = \Storage::disk('s3')->readStream('users.csv');
$excel = \Excel::openStream($stream, ['delimiter' => ';', 'encoding' => 'CP1251']);
$rows = $excel->readRows(true);
fclose($stream);
```

## Temporary files

The temporary copy is created in the directory used for every temporary file of the package: the `temp_dir`
option, then `config('fast-excel.temp_dir')`, then `storage_path('app/tmp/fast-excel')`. It is deleted when
the script ends (and immediately if opening fails), so nothing has to be cleaned up by hand; files left after
a killed process are swept on the next call, like the ones of the writer.

```php
$excel = \Excel::openString($content, ['temp_dir' => '/path/to/tmp']);
```

## Notes and limitations

* **`openStream()` does not close the stream** — the caller keeps its ownership and closes it. The stream is
  copied from its **current position** and is not rewound first, so a stream that has already been read
  yields no data; non-rewindable streams (HTTP wrappers, pipes) work.
* **Memory.** `openStream()` copies the stream to disk in chunks, so a huge workbook never has to fit in
  memory; `openString()` obviously requires the whole content as a string, so prefer the stream form when
  the source can give you one.
* Empty input is rejected: an empty string, a non-resource argument or a stream that produces no data
  throw `\avadim\FastExcelReader\Exception`.
* Writing to a string or a stream is **not** supported. To store an exported workbook use
  [`ExcelWriter::store()`](92-api-class-excelwriter.md) (any Laravel disk) or
  [`ExcelWriter::download()`](92-api-class-excelwriter.md).
