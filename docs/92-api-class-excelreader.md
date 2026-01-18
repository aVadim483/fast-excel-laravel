# Class \avadim\FastExcelLaravel\ExcelReader

---

* [__construct()](#__construct) -- Excel constructor
* [colLetter()](#colletter) -- Convert column number to letter
* [colNum()](#colnum) -- Converts an alphabetic column index to a numeric
* [createReader()](#createreader)
* [createSheet()](#createsheet) -- Create SheetReader instance
* [open()](#open) -- Open XLSX file for import
* [setTempDir()](#settempdir) -- Set dir for temporary files
* [validate()](#validate)
* [countExtraImages()](#countextraimages)
* [countImages()](#countimages) -- Returns the total count of images in the workbook
* [dateFormatter()](#dateformatter) -- Sets custom date formatter
* [formatDate()](#formatdate)
* [from()](#from) -- Set top left of read area
* [getCompleteStyleByIdx()](#getcompletestylebyidx)
* [getDateFormat()](#getdateformat)
* [getDateFormatPattern()](#getdateformatpattern)
* [getDateFormatter()](#getdateformatter)
* [getDefinedNames()](#getdefinednames) -- Returns defined names of workbook
* [getFirstSheet()](#getfirstsheet) -- Returns the first sheet as default
* [getFormatPattern()](#getformatpattern)
* [getImageList()](#getimagelist) -- Returns the list of images from the workbook
* [getSheet()](#getsheet) -- Returns a sheet by name
* [getSheetById()](#getsheetbyid) -- Returns a sheet by ID
* [getSheetNames()](#getsheetnames) -- Returns names array of all sheets
* [hasDrawings()](#hasdrawings) -- Returns TRUE if the workbook contains an any draw objects (not images only)
* [hasExtraImages()](#hasextraimages)
* [hasImages()](#hasimages) -- Returns TRUE if any sheet contains an image object
* [importModel()](#importmodel) -- Import data into a model from the current sheet
* [innerFileList()](#innerfilelist)
* [mapping()](#mapping) -- Set mapping callback for the current sheet
* [mediaImageFiles()](#mediaimagefiles)
* [metadataImage()](#metadataimage)
* [readCallback()](#readcallback) -- Reads cell values and passes them to a callback function
* [readCells()](#readcells) -- Returns the values of all cells as array
* [readCellStyles()](#readcellstyles) -- Returns the styles of all cells as array
* [readCellsWithStyles()](#readcellswithstyles) -- Returns the values and styles of all cells as array
* [readColumns()](#readcolumns) -- Returns cell values as a two-dimensional array from default sheet
* [readColumnsWithStyles()](#readcolumnswithstyles) -- Returns cell values and styles as a two-dimensional array from default sheet
* [readRows()](#readrows) -- Returns cell values as a two-dimensional array from default sheet
* [readRowsWithStyles()](#readrowswithstyles) -- Returns cell values and styles as a two-dimensional array from default sheet
* [readStyles()](#readstyles)
* [selectFirstSheet()](#selectfirstsheet) -- Selects the first sheet as default
* [selectSheet()](#selectsheet) -- Selects default sheet by name
* [selectSheetById()](#selectsheetbyid) -- Selects default sheet by ID
* [setDateFormat()](#setdateformat)
* [setReadArea()](#setreadarea) -- Set top left and right bottom of read area
* [sharedString()](#sharedstring) -- Returns string by index
* [sheet()](#sheet) -- Returns current or specified sheet
* [sheets()](#sheets) -- Array of all sheets
* [styleByIdx()](#stylebyidx) -- Returns a style array by style Idx
* [timestamp()](#timestamp) -- Convert date to timestamp
* [withHeadings()](#withheadings) -- Set headings for the current sheet

---

## __construct()

---

```php
public function __construct(?string $file = null, ?string $tempDir = '')
```
_Excel constructor_

### Parameters

* `string|null $file`
* `string|null $tempDir`

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

## colNum()

---

```php
public static function colNum(string $colLetter): int
```
_Converts an alphabetic column index to a numeric_

### Parameters

* `string $colLetter`

---

## createReader()

---

```php
public static function createReader(string $file, 
                                    ?array $parserProperties = []): avadim\FastExcelReader\Interfaces\InterfaceXmlReader
```


### Parameters

* `string $file`
* `array|null $parserProperties`

---

## createSheet()

---

```php
public static function createSheet(string $sheetName, $sheetId, $file, $path, 
                                   $excel): SheetReader
```
_Create SheetReader instance_

### Parameters

* `string $sheetName`
* `$sheetId`
* `$file`
* `$path`
* `$excel`

---

## open()

---

```php
public static function open(string $file): ExcelReader
```
_Open XLSX file for import_

### Parameters

* `string $file`

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

## validate()

---

```php
public static function validate(string $file, ?array &$errors = []): bool
```


### Parameters

* `string $file`
* `array|null $errors`

---

## countExtraImages()

---

```php
public function countExtraImages(): int
```


### Parameters

_None_

---

## countImages()

---

```php
public function countImages(): int
```
_Returns the total count of images in the workbook_

### Parameters

_None_

---

## dateFormatter()

---

```php
public function dateFormatter($formatter): avadim\FastExcelReader\Excel
```
_Sets custom date formatter_

### Parameters

* `\Closure|callable|string|bool $formatter`

---

## formatDate()

---

```php
public function formatDate($value, $format, $styleIdx): false|mixed|string
```


### Parameters

* `$value`
* `$format`
* `$styleIdx`

---

## from()

---

```php
public function from(string $topLeftCell, 
                     ?bool $firstRowKeys = false): avadim\FastExcelReader\Sheet
```
_Set top left of read area_

### Parameters

* `string $topLeftCell`
* `bool|null $firstRowKeys`

---

## getCompleteStyleByIdx()

---

```php
public function getCompleteStyleByIdx(int $styleIdx, 
                                      ?bool $flat = false): array
```


### Parameters

* `int $styleIdx`
* `bool|null $flat`

---

## getDateFormat()

---

```php
public function getDateFormat(): ?string
```


### Parameters

_None_

---

## getDateFormatPattern()

---

```php
public function getDateFormatPattern(int $styleIdx): ?string
```


### Parameters

* `int $styleIdx`

---

## getDateFormatter()

---

```php
public function getDateFormatter(): callable|\Closure|bool|null
```


### Parameters

_None_

---

## getDefinedNames()

---

```php
public function getDefinedNames(): array
```
_Returns defined names of workbook_

### Parameters

_None_

---

## getFirstSheet()

---

```php
public function getFirstSheet(?string $areaRange = null, 
                              ?bool $firstRowKeys = false): avadim\FastExcelReader\Sheet
```
_Returns the first sheet as default_

### Parameters

* `string|null $areaRange`
* `bool|null $firstRowKeys`

---

## getFormatPattern()

---

```php
public function getFormatPattern(int $styleIdx): mixed|string
```


### Parameters

* `int $styleIdx`

---

## getImageList()

---

```php
public function getImageList(): array
```
_Returns the list of images from the workbook_

### Parameters

_None_

---

## getSheet()

---

```php
public function getSheet(?string $name = null, ?string $areaRange = null, 
                         ?bool $firstRowKeys = false): avadim\FastExcelReader\Sheet
```
_Returns a sheet by name_

### Parameters

* `string|null $name`
* `string|null $areaRange`
* `bool|null $firstRowKeys`

---

## getSheetById()

---

```php
public function getSheetById(int $sheetId, ?string $areaRange = null, 
                             ?bool $firstRowKeys = false): avadim\FastExcelReader\Sheet
```
_Returns a sheet by ID_

### Parameters

* `int $sheetId`
* `string|null $areaRange`
* `bool|null $firstRowKeys`

---

## getSheetNames()

---

```php
public function getSheetNames(): array
```
_Returns names array of all sheets_

### Parameters

_None_

---

## hasDrawings()

---

```php
public function hasDrawings(): bool
```
_Returns TRUE if the workbook contains an any draw objects (not images only)_

### Parameters

_None_

---

## hasExtraImages()

---

```php
public function hasExtraImages(): bool
```


### Parameters

_None_

---

## hasImages()

---

```php
public function hasImages(): bool
```
_Returns TRUE if any sheet contains an image object_

### Parameters

_None_

---

## importModel()

---

```php
public function importModel(string $modelClass, $address, 
                            $columns): ExcelReader
```
_Import data into a model from the current sheet_

### Parameters

* `string $modelClass`
* `string|bool|null $address`
* `array|bool|null $columns`

---

## innerFileList()

---

```php
public function innerFileList(): array
```


### Parameters

_None_

---

## mapping()

---

```php
public function mapping($callback): ExcelReader
```
_Set mapping callback for the current sheet_

### Parameters

* `$callback`

---

## mediaImageFiles()

---

```php
public function mediaImageFiles(): array
```


### Parameters

_None_

---

## metadataImage()

---

```php
public function metadataImage(int $vmIndex): ?string
```


### Parameters

* `int $vmIndex`

---

## readCallback()

---

```php
public function readCallback(callable $callback, ?int $resultMode = null, 
                             ?bool $styleIdxInclude = null)
```
_Reads cell values and passes them to a callback function_

### Parameters

* `callback $callback`
* `int|null $resultMode`
* `bool|null $styleIdxInclude`

---

## readCells()

---

```php
public function readCells(): array
```
_Returns the values of all cells as array_

### Parameters

_None_

---

## readCellStyles()

---

```php
public function readCellStyles(?bool $flat = false): array
```
_Returns the styles of all cells as array_

### Parameters

* `bool|null $flat`

---

## readCellsWithStyles()

---

```php
public function readCellsWithStyles(): array
```
_Returns the values and styles of all cells as array_

### Parameters

_None_

---

## readColumns()

---

```php
public function readColumns($columnKeys, ?int $resultMode = null): array
```
_Returns cell values as a two-dimensional array from default sheet \[col]\[row]_

### Parameters

* `array|bool|int|null $columnKeys`
* `int|null $resultMode`

---

## readColumnsWithStyles()

---

```php
public function readColumnsWithStyles($columnKeys, 
                                      ?int $resultMode = null): array
```
_Returns cell values and styles as a two-dimensional array from default sheet \[col]\[row]_

### Parameters

* `array|bool|int|null $columnKeys`
* `int|null $resultMode`

---

## readRows()

---

```php
public function readRows($columnKeys, ?int $resultMode = null, 
                         ?bool $styleIdxInclude = null): array
```
_Returns cell values as a two-dimensional array from default sheet \[row]\[col]readRows()readRows(true)readRows(false, Excel::KEYS_ZERO_BASED)readRows(Excel::KEYS_ZERO_BASED | Excel::KEYS_RELATIVE)_

### Parameters

* `array|bool|int|null $columnKeys`
* `int|null $resultMode`
* `bool|null $styleIdxInclude`

---

## readRowsWithStyles()

---

```php
public function readRowsWithStyles($columnKeys, 
                                   ?int $resultMode = null): array
```
_Returns cell values and styles as a two-dimensional array from default sheet \[row]\[col]_

### Parameters

* `array|bool|int|null $columnKeys`
* `int|null $resultMode`

---

## readStyles()

---

```php
public function readStyles(): array
```


### Parameters

_None_

---

## selectFirstSheet()

---

```php
public function selectFirstSheet(?string $areaRange = null, 
                                 ?bool $firstRowKeys = false): avadim\FastExcelReader\Sheet
```
_Selects the first sheet as default_

### Parameters

* `string|null $areaRange`
* `bool|null $firstRowKeys`

---

## selectSheet()

---

```php
public function selectSheet(string $name, ?string $areaRange = null, 
                            ?bool $firstRowKeys = false): avadim\FastExcelReader\Sheet
```
_Selects default sheet by name_

### Parameters

* `string $name`
* `string|null $areaRange`
* `bool|null $firstRowKeys`

---

## selectSheetById()

---

```php
public function selectSheetById(int $sheetId, ?string $areaRange = null, 
                                ?bool $firstRowKeys = false): avadim\FastExcelReader\Sheet
```
_Selects default sheet by ID_

### Parameters

* `int $sheetId`
* `string|null $areaRange`
* `bool|null $firstRowKeys`

---

## setDateFormat()

---

```php
public function setDateFormat(string $dateFormat): avadim\FastExcelReader\Excel
```


### Parameters

* `string $dateFormat`

---

## setReadArea()

---

```php
public function setReadArea(string $areaRange, 
                            ?bool $firstRowKeys = false): avadim\FastExcelReader\Sheet
```
_Set top left and right bottom of read area_

### Parameters

* `string $areaRange`
* `bool|null $firstRowKeys`

---

## sharedString()

---

```php
public function sharedString($stringId): ?string
```
_Returns string by index_

### Parameters

* `$stringId`

---

## sheet()

---

```php
public function sheet(?string $name = null): ?avadim\FastExcelReader\Sheet
```
_Returns current or specified sheet_

### Parameters

* `string|null $name`

---

## sheets()

---

```php
public function sheets(): array
```
_Array of all sheets_

### Parameters

_None_

---

## styleByIdx()

---

```php
public function styleByIdx($styleIdx): array
```
_Returns a style array by style Idx_

### Parameters

* `$styleIdx`

---

## timestamp()

---

```php
public function timestamp($excelDateTime): int
```
_Convert date to timestamp_

### Parameters

* `$excelDateTime`

---

## withHeadings()

---

```php
public function withHeadings(?array $headers = []): ExcelReader
```
_Set headings for the current sheet_

### Parameters

* `array|null $headers`

---

