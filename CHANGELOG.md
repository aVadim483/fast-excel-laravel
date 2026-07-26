# Changelog

🇬🇧 English · [🇷🇺 Русский](CHANGELOG.ru.md)

<!-- NOTE for maintainers: this file exists in two languages.
     When editing CHANGELOG.md, please update CHANGELOG.ru.md accordingly.
     Code and identifiers must stay byte-for-byte identical in both versions; only prose is translated. -->

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).
For earlier history see the
[releases page](https://github.com/aVadim483/fast-excel-laravel/releases).

## 4.1.0 - 2026-07-26

### Added

* **Reading of CSV files.** `\Excel::open()` now reads CSV in addition to XLSX and legacy XLS. A file that is
  neither a valid XLSX (ZIP) nor XLS (OLE2) is read as CSV, so the whole import API (`importModel()`,
  `withHeadings()`, `mapping()`, read areas, `readRows()`, `nextRow()`, …) works the same way. CSV values are
  read as strings and CSV carries no styles; writing CSV is not supported (export is still XLSX only). See
  [docs/81-reading-csv.md](docs/81-reading-csv.md).
* **CSV options.** `\Excel::open()` accepts CSV options (`delimiter`, `enclosure`, `escape`, `encoding`,
  `skip_empty_lines`, `comment_prefix`, `mode`) as its second argument, with global defaults in the new `csv`
  section of `config/fast-excel.php`.

### Changed

* Raised the `avadim/fast-excel-reader` requirement to `^4.3` (adds the CSV reader).

## 4.0.0 - 2026-07-25

### Added

* **Reading of legacy XLS workbooks (Excel 97-2003, BIFF8).** `\Excel::open()` now accepts XLS as well as
  XLSX; the format is detected from the file signature, so the extension is ignored and the whole import API
  (`importModel()`, `withHeadings()`, `mapping()`, read areas, `readRows()`, `nextRow()`, …) works the same
  way for both. Writing XLS is not supported — export is still XLSX only. See
  [docs/80-reading-xls.md](docs/80-reading-xls.md).

### Changed

* Raised the `avadim/fast-excel-reader` requirement to `^4.0`.
* **`ExcelReader` and `SheetReader` now use composition instead of inheritance.** They wrap an
  `avadim\FastExcelReader\AbstractBook` / `AbstractSheet` and delegate every call the wrapper does not define,
  so one implementation serves every reader format. The writer classes are unchanged.

### BREAKING

* **`ExcelReader` no longer extends `\avadim\FastExcelReader\Excel`, and `SheetReader` no longer extends
  `\avadim\FastExcelReader\Sheet`.** Objects returned by `\Excel::open()` and `->sheet()` are therefore no
  longer instances of those reader classes.
* Static helpers that used to be inherited on the wrapper (`ExcelReader::validate()`, `::colLetter()`,
  `::colNum()`, `::createReader()`, `::setTempDir()`, …) are gone. Call them on
  `\avadim\FastExcelReader\Excel` directly.
* Class constants that used to be inherited on the wrapper (`ExcelReader::KEYS_*`) are gone. Reference them on
  `\avadim\FastExcelReader\Excel` (that path always worked and is unchanged).
* The internal factory `ExcelReader::createSheet()` was removed (the reader no longer creates the wrapper's
  sheets).

### Upgrade guide (3.x → 4.x)

* Normal usage does not change: `\Excel::open()`, `->withHeadings()`, `->mapping()`, `->importModel()`,
  `->readRows()`, `->from()`, `->setDateFormat()`, `['temp_dir' => …]` and the other documented calls all work
  as before.
* If your code type-hints or checks `instanceof` against `\avadim\FastExcelReader\Excel` / `\Sheet` for values
  coming from this package, switch to `\avadim\FastExcelLaravel\ExcelReader` / `\avadim\FastExcelLaravel\SheetReader`.
* If your code referenced the inherited static helpers or `KEYS_*` constants through the wrapper class, move
  those references to `\avadim\FastExcelReader\Excel`.
