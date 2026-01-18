# Class \avadim\FastExcelLaravel\Excel

---

* [__construct()](#__construct) – Excel constructor
* [create()](#create) – Create new XLSX-file for export
* [open()](#open) – Open an existing XLSX-file for import

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
_Open an existing XLSX-file for import_

### Parameters

* `string $file`
* `array|null $options`

---

