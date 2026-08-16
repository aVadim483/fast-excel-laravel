# Class \avadim\FastExcelLaravel\Excel

<!-- Hand-maintained. Do not auto-generate; edit by hand when the class changes. -->

---

* [__construct()](#__construct) -- Excel constructor
* [create()](#create) -- Create new XLSX-file for export
* [open()](#open) -- Open an existing workbook (XLSX or legacy XLS) for import
* [openString()](#openstring) -- Open a workbook held in a string for import
* [openStream()](#openstream) -- Open a workbook from an open stream resource for import

---

## __construct()

---

```php
public function __construct(?array $options = [])
```
_Excel constructor_

### Parameters

* `array|null $options`

---

## create()

---

```php
public static function create($sheets, ?array $options = []): ExcelWriter
```
_Create new XLSX-file for export_

### Parameters

* `array|string|null $sheets`
* `array|null $options`

---

## open()

---

```php
public static function open(string $file, ?array $options = []): ExcelReader
```
_Open an existing workbook (XLSX or legacy XLS) for import. The format is detected automatically from the file
signature, so the file extension does not matter._

### Parameters

* `string $file`
* `array|null $options`

---

## openString()

---

```php
public static function openString(string $content, ?array $options = []): ExcelReader
```
_Open a workbook held in a string (a database blob, an HTTP response body, `Storage::get()`, …). The content
is written to a temporary file and opened like [open()](#open), so the format is detected from the bytes and
the same options apply. See [Reading from a string or a stream](82-reading-from-memory.md)._

### Parameters

* `string $content`
* `array|null $options` -- same as [open()](#open)

---

## openStream()

---

```php
public static function openStream($stream, ?array $options = []): ExcelReader
```
_Open a workbook from an open readable stream resource, e.g. `\Storage::disk('s3')->readStream($path)`. The
stream is copied into a temporary file from its current position and opened like [open()](#open); it is not
closed. See [Reading from a string or a stream](82-reading-from-memory.md)._

### Parameters

* `resource $stream`
* `array|null $options` -- same as [open()](#open)

---

