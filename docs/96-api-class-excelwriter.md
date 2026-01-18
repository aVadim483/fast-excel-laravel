# Class \avadim\FastExcelLaravel\ExcelWriter

---

* [__construct()](#__construct) -- Excel constructor
* [cellAddress()](#celladdress) -- Create cell address by row and col numbers
* [colIndex()](#colindex) -- Convert letter to index (ZERO based)
* [colIndexRange()](#colindexrange) -- Convert letter range to array of numbers (ZERO based)
* [colLetter()](#colletter) -- Convert column number to letter
* [colLetterRange()](#colletterrange) -- Convert values to letters array
* [colNumber()](#colnumber) -- Convert letter to number (ONE based)
* [colNumberRange()](#colnumberrange) -- Convert letter range to array of numbers (ONE based)
* [create()](#create) -- Create XLSX for export
* [createSheet()](#createsheet) -- Create SheetWriter instance
* [fullAddress()](#fulladdress) -- Get full address with sheet name
* [generateUuid()](#generateuuid) -- Generate UUID v4
* [newStyle()](#newstyle) -- Create a new instance of Style class
* [pixelsToEMU()](#pixelstoemu) -- Convert pixels to EMU (English Metric Units)
* [rangeRelOffsets()](#rangereloffsets) -- Return offsets by relative address (zero based)
* [setTempDir()](#settempdir) -- Set dir for temporary files
* [toTimestamp()](#totimestamp) -- Convert value (int or string) to Excel timestamp
* [addDefinedName()](#adddefinedname) -- Add a defined name
* [addNamedRange()](#addnamedrange) -- Add a named range
* [addSharedString()](#addsharedstring) -- Add a string to the shared strings table
* [download()](#download) -- Download generated file to client (send to browser)
* [exportModel()](#exportmodel) -- Export a model to the current sheet
* [getDefaultFormatStyles()](#getdefaultformatstyles) -- Get default format styles
* [getDefaultSheetName()](#getdefaultsheetname) -- Get the default sheet name
* [getDefaultStyle()](#getdefaultstyle) -- Get default style
* [getDefinedNames()](#getdefinednames) -- Get all defined names
* [getFileName()](#getfilename) -- Returns default filename
* [getImageFiles()](#getimagefiles) -- Get list of all loaded image files
* [getMetadata()](#getmetadata) -- Get metadata
* [getSharedStrings()](#getsharedstrings) -- Get the shared strings table
* [getSheet()](#getsheet) -- Alias of sheet()
* [getSheets()](#getsheets) -- Returns all sheets
* [getWriter()](#getwriter) -- Get the writer instance
* [isR1C1()](#isr1c1) -- Is R1C1 notation mode enabled?
* [isRightToLeft()](#isrighttoleft) -- Get right-to-left orientation
* [makeSheet()](#makesheet) -- Creates and returns a new sheet
* [output()](#output) -- Alias of download()
* [protect()](#protect) -- Protect workbook
* [removeSheet()](#removesheet) -- Removes sheet by index or name of sheet.
* [save()](#save) -- Save generated XLSX-file
* [saveTo()](#saveto) -- Save file to local storage
* [setActiveSheet()](#setactivesheet) -- Set active (default) sheet by case-insensitive name
* [setAuthor()](#setauthor) -- Set the author of the document
* [setCompany()](#setcompany) -- Set the company of the document
* [setDefaultFont()](#setdefaultfont) -- Set default font options
* [setDefaultFontName()](#setdefaultfontname) -- Set default font name
* [setDefaultLocale()](#setdefaultlocale) -- Set default locale from the current environment
* [setDefaultStyle()](#setdefaultstyle) -- Set default style
* [setDescription()](#setdescription) -- Set the description of the document
* [setFileName()](#setfilename) -- Sets default filename for saving
* [setKeywords()](#setkeywords) -- Set the keywords of the document
* [setLocale()](#setlocale) -- Set locale information
* [setMetaAuthor()](#setmetaauthor) -- Set metadata 'author'
* [setMetaCompany()](#setmetacompany) -- Set metadata 'company'
* [setMetadata()](#setmetadata) -- Set metadata
* [setMetaDescription()](#setmetadescription) -- Set metadata 'description'
* [setMetaKeywords()](#setmetakeywords) -- Set metadata 'keywords'
* [setMetaSubject()](#setmetasubject) -- Set metadata 'subject'
* [setMetaTitle()](#setmetatitle) -- Set metadata 'title'
* [setR1C1()](#setr1c1) -- Set R1C1 notation mode
* [setRightToLeft()](#setrighttoleft) -- Set right-to-left orientation
* [setSharedString()](#setsharedstring) -- Set the usage of shared strings
* [setSubject()](#setsubject) -- Set the subject of the document
* [setTitle()](#settitle) -- Set the title of the document
* [sheet()](#sheet) -- Returns sheet by number or name of sheet.
* [store()](#store) -- Store file to specified disk
* [unprotect()](#unprotect) -- Unprotect workbook
* [writeData()](#writedata) -- Write data to the current sheet

---

## __construct()

---

```php
public function __construct(?array $options = [])
```
_Excel constructor_

### Parameters

* `array|null $options` -- Optional parameters: \['temp_dir' => ..., 'temp_prefix' => ..., 'auto_convert_number' => ..., 'shared_string' => ...]

---

## cellAddress()

---

```php
public static function cellAddress(int $rowNumber, int $colNumber, 
                                   ?bool $absolute = false, 
                                   ?bool $absoluteRow = null): string
```
_Create cell address by row and col numbers_

### Parameters

* `int $rowNumber` -- ONE based
* `int $colNumber` -- ONE based
* `bool|null $absolute`
* `bool|null $absoluteRow`

---

### Examples

```php
cellAddress(3, 3) => 'C3'
cellAddress(43, 27) => 'AA43'
cellAddress(43, 27, true) => '$AA$43'
cellAddress(43, 27, false, true) => 'AA$43'
```


---

## colIndex()

---

```php
public static function colIndex(string $colLetter): int
```
_Convert letter to index (ZERO based)_

### Parameters

* `string $colLetter`

---

## colIndexRange()

---

```php
public static function colIndexRange($colLetter): array
```
_Convert letter range to array of numbers (ZERO based)_

### Parameters

* `string|int|array $colLetter` -- e.g.: 'B', 2, 'C:F', \['A', 'B', 'C']

---

## colLetter()

---

```php
public static function colLetter(int $colNumber): string
```
_Convert column number to letter_

### Parameters

* `int $colNumber` -- ONE based

---

## colLetterRange()

---

```php
public static function colLetterRange($colKeys, ?int $baseNum = 0): array
```
_Convert values to letters array_

### Parameters

* `array|string $colKeys`
* `int|null $baseNum` -- 0 or 1

---

### Examples

```php
$res = colLetterRange([0, 1, 2]);    // returns ['A', 'B', 'C']
$res = colLetterRange([1, 2, 3], 1); // returns ['A', 'B', 'C']
$res = colLetterRange('B, E, F');    // returns ['B', 'E', 'F']
$res = colLetterRange('B-E, F');     // returns ['B', 'C', 'D', 'E', 'F']
$res = colLetterRange('B1-E8');      // returns ['B', 'C', 'D', 'E']
$res = colLetterRange('B1:E8');      // returns ['B:E']
```


---

## colNumber()

---

```php
public static function colNumber(string $colLetter): int
```
_Convert letter to number (ONE based)_

### Parameters

* `string $colLetter`

---

## colNumberRange()

---

```php
public static function colNumberRange($colLetter): array
```
_Convert letter range to array of numbers (ONE based)_

### Parameters

* `string|int|array $colLetter` -- e.g.: 'B', 2, 'C:F', \['A', 'B', 'C']

---

## create()

---

```php
public static function create($sheets, ?array $options = []): ExcelWriter
```
_Create XLSX for export_

### Parameters

* `string|array $sheets`
* `array|null $options`

---

## createSheet()

---

```php
public static function createSheet(string $sheetName): SheetWriter
```
_Create SheetWriter instance_

### Parameters

* `string $sheetName`

---

## fullAddress()

---

```php
public static function fullAddress(string $sheetName, string $address, 
                                   ?bool $force = false): string
```
_Get full address with sheet name_

### Parameters

* `string $sheetName`
* `string $address`
* `bool|null $force`

---

## generateUuid()

---

```php
public static function generateUuid(): string
```
_Generate UUID v4_

### Parameters

_None_

---

## newStyle()

---

```php
public static function newStyle(): avadim\FastExcelWriter\Style\Style
```
_Create a new instance of Style class_

### Parameters

_None_

---

## pixelsToEMU()

---

```php
public static function pixelsToEMU($pixels): float|int
```
_Convert pixels to EMU (English Metric Units)_

### Parameters

* `int|float $pixels`

---

## rangeRelOffsets()

---

```php
public static function rangeRelOffsets(string $relAddress): array
```
_Return offsets by relative address (zero based)_

### Parameters

* `string $relAddress`

---

## setTempDir()

---

```php
public static function setTempDir($tempDir)
```
_Set dir for temporary files_

### Parameters

* `$tempDir`

---

## toTimestamp()

---

```php
public static function toTimestamp($value): float|bool
```
_Convert value (int or string) to Excel timestamp_

### Parameters

* `int|string $value`

---

## addDefinedName()

---

```php
public function addDefinedName(string $name, string $range, 
                               ?array $attributes = []): avadim\FastExcelWriter\Excel
```
_Add a defined name_

### Parameters

* `string $name`
* `string $range`
* `array|null $attributes`

---

## addNamedRange()

---

```php
public function addNamedRange(string $range, 
                              string $name): avadim\FastExcelWriter\Excel
```
_Add a named range_

### Parameters

* `string $range`
* `string $name`

---

## addSharedString()

---

```php
public function addSharedString(string $string, ?bool $richText = false): int
```
_Add a string to the shared strings table_

### Parameters

* `string $string`
* `bool|null $richText`

---

## download()

---

```php
public function download(?string $name = null)
```
_Download generated file to client (send to browser)_

### Parameters

* `string|null $name`

---

## exportModel()

---

```php
public function exportModel($model, ?array $rowStyle = null, 
                            ?array $cellStyles = null): ExcelWriter
```
_Export a model to the current sheet_

### Parameters

* `$model`
* `array|null $rowStyle`
* `array|null $cellStyles`

---

## getDefaultFormatStyles()

---

```php
public function getDefaultFormatStyles(): array
```
_Get default format styles_

### Parameters

_None_

---

## getDefaultSheetName()

---

```php
public function getDefaultSheetName(): string
```
_Get the default sheet name_

### Parameters

_None_

---

## getDefaultStyle()

---

```php
public function getDefaultStyle(): array
```
_Get default style_

### Parameters

_None_

---

## getDefinedNames()

---

```php
public function getDefinedNames(): array
```
_Get all defined names_

### Parameters

_None_

---

## getFileName()

---

```php
public function getFileName(): string
```
_Returns default filename_

### Parameters

_None_

---

## getImageFiles()

---

```php
public function getImageFiles(): array
```
_Get list of all loaded image files_

### Parameters

_None_

---

## getMetadata()

---

```php
public function getMetadata($key): mixed
```
_Get metadata_

### Parameters

* `null $key`

---

## getSharedStrings()

---

```php
public function getSharedStrings(): array
```
_Get the shared strings table_

### Parameters

_None_

---

## getSheet()

---

```php
public function getSheet($index): ?avadim\FastExcelWriter\Sheet
```
_Alias of sheet()_

### Parameters

* `int|string|null $index` -- - number or name of sheet

---

## getSheets()

---

```php
public function getSheets(): array
```
_Returns all sheets_

### Parameters

_None_

---

## getWriter()

---

```php
public function getWriter(): Writer
```
_Get the writer instance_

### Parameters

_None_

---

## isR1C1()

---

```php
public function isR1C1(): bool
```
_Is R1C1 notation mode enabled?_

### Parameters

_None_

---

## isRightToLeft()

---

```php
public function isRightToLeft(): bool
```
_Get right-to-left orientation_

### Parameters

_None_

---

## makeSheet()

---

```php
public function makeSheet(?string $sheetName = null): avadim\FastExcelWriter\Sheet
```
_Creates and returns a new sheet_

### Parameters

* `string|null $sheetName`

---

## output()

---

```php
public function output(?string $name = null): void
```
_Alias of download()_

### Parameters

* `string|null $name`

---

## protect()

---

```php
public function protect(?string $password = null): avadim\FastExcelWriter\Excel
```
_Protect workbook_

### Parameters

* `string|null $password`

---

## removeSheet()

---

```php
public function removeSheet($index): avadim\FastExcelWriter\Excel
```
_Removes sheet by index or name of sheet.Removes the first sheet of index omitted_

### Parameters

* `int|string|null $index`

---

## save()

---

```php
public function save(?string $fileName = null, ?bool $overWrite = true): bool
```
_Save generated XLSX-file_

### Parameters

* `string|null $fileName`
* `bool|null $overWrite`

---

## saveTo()

---

```php
public function saveTo(string $filePath): bool
```
_Save file to local storage_

### Parameters

* `string $filePath`

---

## setActiveSheet()

---

```php
public function setActiveSheet(string $name): avadim\FastExcelWriter\Excel
```
_Set active (default) sheet by case-insensitive name_

### Parameters

* `string $name`

---

## setAuthor()

---

```php
public function setAuthor(?string $author = ''): avadim\FastExcelWriter\Excel
```
_Set the author of the document_

### Parameters

* `string|null $author`

---

## setCompany()

---

```php
public function setCompany(?string $company = ''): avadim\FastExcelWriter\Excel
```
_Set the company of the document_

### Parameters

* `string|null $company`

---

## setDefaultFont()

---

```php
public function setDefaultFont(array $fontOptions): avadim\FastExcelWriter\Excel
```
_Set default font options_

### Parameters

* `array $fontOptions`

---

## setDefaultFontName()

---

```php
public function setDefaultFontName(string $fontName): avadim\FastExcelWriter\Excel
```
_Set default font name_

### Parameters

* `string $fontName`

---

## setDefaultLocale()

---

```php
public function setDefaultLocale()
```
_Set default locale from the current environment_

### Parameters

_None_

---

## setDefaultStyle()

---

```php
public function setDefaultStyle(array $style): avadim\FastExcelWriter\Excel
```
_Set default style_

### Parameters

* `array $style`

---

## setDescription()

---

```php
public function setDescription(?string $description = ''): avadim\FastExcelWriter\Excel
```
_Set the description of the document_

### Parameters

* `string|null $description`

---

## setFileName()

---

```php
public function setFileName(string $fileName): avadim\FastExcelWriter\Excel
```
_Sets default filename for saving_

### Parameters

* `string $fileName`

---

## setKeywords()

---

```php
public function setKeywords($keywords): avadim\FastExcelWriter\Excel
```
_Set the keywords of the document_

### Parameters

* `string|array $keywords`

---

## setLocale()

---

```php
public function setLocale(string $locale, 
                          ?string $dir = null): avadim\FastExcelWriter\Excel
```
_Set locale information_

### Parameters

* `string $locale`
* `string|null $dir`

---

## setMetaAuthor()

---

```php
public function setMetaAuthor(?string $author = ''): avadim\FastExcelWriter\Excel
```
_Set metadata 'author'_

### Parameters

* `string|null $author`

---

## setMetaCompany()

---

```php
public function setMetaCompany(?string $company = ''): avadim\FastExcelWriter\Excel
```
_Set metadata 'company'_

### Parameters

* `string|null $company`

---

## setMetadata()

---

```php
public function setMetadata($key, $value): avadim\FastExcelWriter\Excel
```
_Set metadata_

### Parameters

* `$key`
* `$value`

---

## setMetaDescription()

---

```php
public function setMetaDescription(?string $description = ''): avadim\FastExcelWriter\Excel
```
_Set metadata 'description'_

### Parameters

* `string|null $description`

---

## setMetaKeywords()

---

```php
public function setMetaKeywords($keywords): avadim\FastExcelWriter\Excel
```
_Set metadata 'keywords'_

### Parameters

* `mixed $keywords`

---

## setMetaSubject()

---

```php
public function setMetaSubject(?string $subject = ''): avadim\FastExcelWriter\Excel
```
_Set metadata 'subject'_

### Parameters

* `string|null $subject`

---

## setMetaTitle()

---

```php
public function setMetaTitle(?string $title = ''): avadim\FastExcelWriter\Excel
```
_Set metadata 'title'_

### Parameters

* `string|null $title`

---

## setR1C1()

---

```php
public function setR1C1(bool $option = true): avadim\FastExcelWriter\Excel
```
_Set R1C1 notation mode_

### Parameters

* `bool $option`

---

## setRightToLeft()

---

```php
public function setRightToLeft(bool $isRightToLeft)
```
_Set right-to-left orientation_

### Parameters

* `bool $isRightToLeft`

---

## setSharedString()

---

```php
public function setSharedString(bool $option = true): avadim\FastExcelWriter\Excel
```
_Set the usage of shared strings_

### Parameters

* `bool $option`

---

## setSubject()

---

```php
public function setSubject(?string $subject = ''): avadim\FastExcelWriter\Excel
```
_Set the subject of the document_

### Parameters

* `string|null $subject`

---

## setTitle()

---

```php
public function setTitle(?string $title = ''): avadim\FastExcelWriter\Excel
```
_Set the title of the document_

### Parameters

* `string|null $title`

---

## sheet()

---

```php
public function sheet($index): SheetWriter
```
_Returns sheet by number or name of sheet.Return the first sheet if number or name omitted_

### Parameters

* `int|string|null $index` -- - number or name of sheet

---

## store()

---

```php
public function store($disk, $path): bool
```
_Store file to specified disk_

### Parameters

* `$disk`
* `$path`

---

## unprotect()

---

```php
public function unprotect(): avadim\FastExcelWriter\Excel
```
_Unprotect workbook_

### Parameters

_None_

---

## writeData()

---

```php
public function writeData($data): ExcelWriter
```
_Write data to the current sheet_

### Parameters

* `$data`

---

