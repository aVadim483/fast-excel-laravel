# Class \avadim\FastExcelLaravel\Excel

---

* [__construct()](#__construct) – Excel constructor
* [create()](#create) – Create new XLSX-file for export
* [open()](#open) – Open an existing workbook (XLSX or legacy XLS) for import

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

