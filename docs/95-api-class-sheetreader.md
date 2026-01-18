# Class \avadim\FastExcelLaravel\SheetReader

---

* [__construct()](#__construct)
* [actualDimension()](#actualdimension)
* [countActualColumns()](#countactualcolumns) -- Returns the actual number of columns from the sheet data area
* [countActualDimension()](#countactualdimension)
* [countActualRows()](#countactualrows) -- Returns the actual number of rows from the sheet data area
* [countCols()](#countcols) -- Count columns by dimension value, alias of countColumns()
* [countColumns()](#countcolumns) -- Count columns by dimension value
* [countImages()](#countimages) -- Count images of the sheet
* [countRows()](#countrows) -- Count rows by dimension value
* [dimension()](#dimension)
* [dimensionArray()](#dimensionarray)
* [extractConditionalFormatting()](#extractconditionalformatting) -- Extracts conditional formatting rules from the sheet
* [extractDataValidations()](#extractdatavalidations) -- Extracts data validation rules from the sheet
* [firstCol()](#firstcol)
* [firstRow()](#firstrow)
* [from()](#from) -- Set top left of read area
* [getColAttributes()](#getcolattributes)
* [getColumnAttributes()](#getcolumnattributes)
* [getColumnStyle()](#getcolumnstyle)
* [getColumnWidth()](#getcolumnwidth) -- Returns column width for a specific column number.
* [getConditionalFormatting()](#getconditionalformatting) -- Returns an array of data validation rules found in the sheet
* [getDataValidations()](#getdatavalidations) -- Returns an array of data validation rules found in the sheet
* [getFreezePaneInfo()](#getfreezepaneinfo) -- Parses and retrieves frozen pane info from the sheet XML
* [getImageBlob()](#getimageblob) -- Returns an image from the cell as a blob (if exists) or null
* [getImageList()](#getimagelist)
* [getImageListByRow()](#getimagelistbyrow)
* [getImageMimeType()](#getimagemimetype) -- Returns the MIME type for an image from the cell as determined by using information from the magic.mime file
* [getImageName()](#getimagename) -- Returns the name for an image from the cell as it defines in XLSX
* [getMergedCells()](#getmergedcells) -- Returns all merged ranges
* [getReadRowNum()](#getreadrownum)
* [getRowHeight()](#getrowheight) -- Returns row height for a specific row number.
* [getTabColorConfiguration()](#gettabcolorconfiguration) -- Alias of getTabColorConfig()
* [getTabColorInfo()](#gettabcolorinfo) -- Returns the tab color info of the sheet
* [hasDrawings()](#hasdrawings)
* [hasImage()](#hasimage) -- Returns TRUE if the cell contains an image
* [id()](#id)
* [imageEntryFullPath()](#imageentryfullpath) -- Returns full path of an image from the cell (if exists) or null
* [importModel()](#importmodel) -- Load models from Excel to database
* [isActive()](#isactive)
* [isHidden()](#ishidden)
* [isMerged()](#ismerged) -- Checks if a cell is merged
* [isName()](#isname) -- Case-insensitive name checking
* [isVisible()](#isvisible)
* [mapping()](#mapping) -- Set mapping callback for the sheet
* [maxActualColumn()](#maxactualcolumn)
* [maxActualRow()](#maxactualrow)
* [maxColumn()](#maxcolumn) -- Max column from dimension value
* [maxRow()](#maxrow) -- Max row number from dimension value
* [mergedRange()](#mergedrange) -- Returns merge range of specified cell
* [minActualColumn()](#minactualcolumn)
* [minActualRow()](#minactualrow)
* [minColumn()](#mincolumn) -- Min column from dimension value
* [minRow()](#minrow) -- Min row number from dimension value
* [name()](#name)
* [nextRow()](#nextrow) -- Read cell values row by row, returns either an array of values or an array of arrays
* [path()](#path)
* [readCallback()](#readcallback) -- Reads cell values and passes them to a callback function
* [readCells()](#readcells) -- Returns values and styles of cells as array
* [readCellStyles()](#readcellstyles) -- Returns styles of cells as array
* [readCellsWithStyles()](#readcellswithstyles) -- Returns values and styles of cells as array:
* [readColumns()](#readcolumns) -- Returns cell values as a two-dimensional array from default sheet
* [readColumnsWithStyles()](#readcolumnswithstyles) -- Returns values and styles of cells as array
* [readFirstRow()](#readfirstrow) -- Returns values of cells of 1st row as array
* [readFirstRowCells()](#readfirstrowcells) -- Returns values and styles of cells of 1st row as array
* [readFirstRowWithStyles()](#readfirstrowwithstyles)
* [readNextRow()](#readnextrow)
* [readRows()](#readrows) -- Returns cell values as a two-dimensional array
* [readRowsWithStyles()](#readrowswithstyles) -- Returns values, styles and other info of cells as array
* [reset()](#reset) -- Reset read generator
* [rewind()](#rewind) -- Rewind read generator, alias of reset()
* [saveImage()](#saveimage) -- Writes an image from the cell to the specified filename
* [saveImageTo()](#saveimageto) -- Writes an image from the cell to the specified directory
* [setDateFormat()](#setdateformat)
* [setDefaultRowHeight()](#setdefaultrowheight)
* [setReadArea()](#setreadarea) -- Set top left and right bottom of read area
* [setReadAreaColumns()](#setreadareacolumns) -- setReadArea('C:AZ') - set left and right columns of read area
* [setState()](#setstate)
* [state()](#state)
* [withHeadings()](#withheadings) -- Set headings for the sheet

---

## __construct()

---

```php
public function __construct(string $sheetName, string $sheetId, string $file, 
                            string $path, $excel)
```


### Parameters

* `string $sheetName`
* `string $sheetId`
* `string $file`
* `string $path`
* `$excel`

---

## actualDimension()

---

```php
public function actualDimension(): string
```


### Parameters

_None_

---

## countActualColumns()

---

```php
public function countActualColumns(): int
```
_Returns the actual number of columns from the sheet data area_

### Parameters

_None_

---

## countActualDimension()

---

```php
public function countActualDimension(bool $countColumns = true, 
                                     bool $countRows = true, 
                                     int $blockSize = 4096): array
```


### Parameters

* `bool $countColumns`
* `bool $countRows`
* `int $blockSize`

---

## countActualRows()

---

```php
public function countActualRows(): int
```
_Returns the actual number of rows from the sheet data area_

### Parameters

_None_

---

## countCols()

---

```php
public function countCols(?string $range = null): int
```
_Count columns by dimension value, alias of countColumns()_

### Parameters

* `string|null $range`

---

## countColumns()

---

```php
public function countColumns(?string $range = null): int
```
_Count columns by dimension value_

### Parameters

* `string|null $range`

---

## countImages()

---

```php
public function countImages(): int
```
_Count images of the sheet_

### Parameters

_None_

---

## countRows()

---

```php
public function countRows(?string $range = null): int
```
_Count rows by dimension value_

### Parameters

* `string|null $range`

---

## dimension()

---

```php
public function dimension(): ?string
```


### Parameters

_None_

---

## dimensionArray()

---

```php
public function dimensionArray(): array
```


### Parameters

_None_

---

## extractConditionalFormatting()

---

```php
public function extractConditionalFormatting(): void
```
_Extracts conditional formatting rules from the sheet_

### Parameters

_None_

---

## extractDataValidations()

---

```php
public function extractDataValidations(): void
```
_Extracts data validation rules from the sheet_

### Parameters

_None_

---

## firstCol()

---

```php
public function firstCol(): string
```


### Parameters

_None_

---

## firstRow()

---

```php
public function firstRow(): int
```


### Parameters

_None_

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

## getColAttributes()

---

```php
public function getColAttributes(): array
```


### Parameters

_None_

---

## getColumnAttributes()

---

```php
public function getColumnAttributes($col): array|mixed
```


### Parameters

* `int|string $col`

---

## getColumnStyle()

---

```php
public function getColumnStyle($col, ?bool $flat = false): array
```


### Parameters

* `int|string $col`
* `bool|null $flat`

---

## getColumnWidth()

---

```php
public function getColumnWidth(int $colNumber): ?float
```
_Returns column width for a specific column number._

### Parameters

* `int $colNumber`

---

## getConditionalFormatting()

---

```php
public function getConditionalFormatting(): array
```
_Returns an array of data validation rules found in the sheet_

### Parameters

_None_

---

## getDataValidations()

---

```php
public function getDataValidations(): array
```
_Returns an array of data validation rules found in the sheet_

### Parameters

_None_

---

## getFreezePaneInfo()

---

```php
public function getFreezePaneInfo(): ?array
```
_Parses and retrieves frozen pane info from the sheet XML_

### Parameters

_None_

---

## getImageBlob()

---

```php
public function getImageBlob(string $cell): ?string
```
_Returns an image from the cell as a blob (if exists) or null_

### Parameters

* `string $cell`

---

## getImageList()

---

```php
public function getImageList(): array
```


### Parameters

_None_

---

## getImageListByRow()

---

```php
public function getImageListByRow($row): array
```


### Parameters

* `$row`

---

## getImageMimeType()

---

```php
public function getImageMimeType(string $cell): ?string
```
_Returns the MIME type for an image from the cell as determined by using information from the magic.mime fileRequires fileinfo extension_

### Parameters

* `string $cell`

---

## getImageName()

---

```php
public function getImageName(string $cell): ?string
```
_Returns the name for an image from the cell as it defines in XLSX_

### Parameters

* `string $cell`

---

## getMergedCells()

---

```php
public function getMergedCells(): ?array
```
_Returns all merged ranges_

### Parameters

_None_

---

## getReadRowNum()

---

```php
public function getReadRowNum(): int
```


### Parameters

_None_

---

## getRowHeight()

---

```php
public function getRowHeight(int $rowNumber): ?float
```
_Returns row height for a specific row number._

### Parameters

* `int $rowNumber`

---

## getTabColorConfiguration()

---

```php
public function getTabColorConfiguration(): ?array
```
_Alias of getTabColorConfig()_

### Parameters

_None_

---

## getTabColorInfo()

---

```php
public function getTabColorInfo(): ?array
```
_Returns the tab color info of the sheetContains any of: rgb, theme, tint, indexed_

### Parameters

_None_

---

## hasDrawings()

---

```php
public function hasDrawings(): bool
```


### Parameters

_None_

---

## hasImage()

---

```php
public function hasImage(string $cell): bool
```
_Returns TRUE if the cell contains an image_

### Parameters

* `string $cell`

---

## id()

---

```php
public function id(): string
```


### Parameters

_None_

---

## imageEntryFullPath()

---

```php
public function imageEntryFullPath(string $cell): ?string
```
_Returns full path of an image from the cell (if exists) or null_

### Parameters

* `string $cell`

---

## importModel()

---

```php
public function importModel($modelClass, $address, $columns): SheetReader
```
_Load models from Excel to databaseloadModels(User::class)loadModels(User::class, true) -- the first row used as a field namesloadModels(User::class, 'B:D') -- read data from columns B:DloadModels(User::class, 'B3') -- read data from area started at B3loadModels(User::class, 'B3', true) -- read data from area started at B3 and the first row used as a field names_

### Parameters

* `$modelClass`
* `$address`
* `$columns`

---

## isActive()

---

```php
public function isActive(): bool
```


### Parameters

_None_

---

## isHidden()

---

```php
public function isHidden(): bool
```


### Parameters

_None_

---

## isMerged()

---

```php
public function isMerged(string $cellAddress): bool
```
_Checks if a cell is merged_

### Parameters

* `string $cellAddress`

---

## isName()

---

```php
public function isName(string $name): bool
```
_Case-insensitive name checking_

### Parameters

* `string $name`

---

## isVisible()

---

```php
public function isVisible(): bool
```


### Parameters

_None_

---

## mapping()

---

```php
public function mapping($callback): SheetReader
```
_Set mapping callback for the sheet_

### Parameters

* `$callback`

---

## maxActualColumn()

---

```php
public function maxActualColumn(): string
```


### Parameters

_None_

---

## maxActualRow()

---

```php
public function maxActualRow(): int
```


### Parameters

_None_

---

## maxColumn()

---

```php
public function maxColumn(?string $range = null): string
```
_Max column from dimension value_

### Parameters

* `string|null $range`

---

## maxRow()

---

```php
public function maxRow(?string $range = null): int
```
_Max row number from dimension value_

### Parameters

* `string|null $range`

---

## mergedRange()

---

```php
public function mergedRange(string $cellAddress): ?string
```
_Returns merge range of specified cell_

### Parameters

* `string $cellAddress`

---

## minActualColumn()

---

```php
public function minActualColumn(): string
```


### Parameters

_None_

---

## minActualRow()

---

```php
public function minActualRow(): int
```


### Parameters

_None_

---

## minColumn()

---

```php
public function minColumn(?string $range = null): string
```
_Min column from dimension value_

### Parameters

* `string|null $range`

---

## minRow()

---

```php
public function minRow(?string $range = null): int
```
_Min row number from dimension value_

### Parameters

* `string|null $range`

---

## name()

---

```php
public function name(): string
```


### Parameters

_None_

---

## nextRow()

---

```php
public function nextRow($columnKeys, ?int $resultMode = null, 
                        ?bool $styleIdxInclude = null, 
                        ?int $rowLimit = 0): ?Generator
```
_Read cell values row by row, returns either an array of values or an array of arraysnextRow(..., ...) : <rowNum> => \[<colNum1> => <value1>, <colNum2> => <value2>, ...]nextRow(..., ..., true) : <rowNum> => \[<colNum1> => \['v' => <value1>, 's' => <style1>], <colNum2> => \['v' => <value2>, 's' => <style2>], ...]_

### Parameters

* `array|bool|int|null $columnKeys`
* `int|null $resultMode`
* `bool|null $styleIdxInclude`
* `int|null $rowLimit`

---

## path()

---

```php
public function path(): string
```


### Parameters

_None_

---

## readCallback()

---

```php
public function readCallback(callable $callback, $columnKeys, 
                             ?int $resultMode = null, 
                             ?bool $styleIdxInclude = null)
```
_Reads cell values and passes them to a callback function_

### Parameters

* `callback $callback` -- Callback function($row, $col, $value)
* `array|bool|int|null $columnKeys`
* `int|null $resultMode`
* `bool|null $styleIdxInclude`

---

## readCells()

---

```php
public function readCells(?bool $styleIdxInclude = null): array
```
_Returns values and styles of cells as array_

### Parameters

* `bool|null $styleIdxInclude`

---

## readCellStyles()

---

```php
public function readCellStyles(?bool $flat = false, 
                               ?string $part = null): array
```
_Returns styles of cells as array_

### Parameters

* `bool|null $flat`
* `string|null $part`

---

## readCellsWithStyles()

---

```php
public function readCellsWithStyles($styleKey): array
```
_Returns values and styles of cells as array:'v' => _value_'s' => _styles_'f' => _formula_'t' => _type_'o' => _original_value__

### Parameters

* `$styleKey`

---

## readColumns()

---

```php
public function readColumns($columnKeys, ?int $resultMode = null, 
                            ?bool $styleIdxInclude = null): array
```
_Returns cell values as a two-dimensional array from default sheet \[col]\[row]\['A' => \[1 => _value_A1_], \[2 => _value_A2_]],\['B' => \[1 => _value_B1_], \[2 => _value_B2_]]_

### Parameters

* `array|bool|int|null $columnKeys`
* `int|null $resultMode`
* `bool|null $styleIdxInclude`

---

## readColumnsWithStyles()

---

```php
public function readColumnsWithStyles($columnKeys, 
                                      ?int $resultMode = null): array
```
_Returns values and styles of cells as array \['v' => _value_, 's' => _styles_]_

### Parameters

* `array|bool|int|null $columnKeys`
* `int|null $resultMode`

---

## readFirstRow()

---

```php
public function readFirstRow($columnKeys, 
                             ?bool $styleIdxInclude = null): array
```
_Returns values of cells of 1st row as array_

### Parameters

* `array|bool|int|null $columnKeys`
* `bool|null $styleIdxInclude`

---

## readFirstRowCells()

---

```php
public function readFirstRowCells(?bool $styleIdxInclude = null): array
```
_Returns values and styles of cells of 1st row as array_

### Parameters

* `bool|null $styleIdxInclude`

---

## readFirstRowWithStyles()

---

```php
public function readFirstRowWithStyles($columnKeys): array
```


### Parameters

* `array|bool|int|null $columnKeys`

---

## readNextRow()

---

```php
public function readNextRow(): mixed
```


### Parameters

_None_

---

## readRows()

---

```php
public function readRows($columnKeys, ?int $resultMode = null, 
                         ?bool $styleIdxInclude = null): array
```
_Returns cell values as a two-dimensional array\[1 => \['A' => _value_A1_], \['B' => _value_B1_]],\[2 => \['A' => _value_A2_], \['B' => _value_B2_]]readRows()readRows(true)readRows(false, Excel::KEYS_ZERO_BASED)readRows(Excel::KEYS_ZERO_BASED | Excel::KEYS_RELATIVE)_

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
_Returns values, styles and other info of cells as array\['v' => _value_,'s' => _styles_,'f' => _formula_,'t' => _type_,'o' => '_original_value_]_

### Parameters

* `array|bool|int|null $columnKeys`
* `int|null $resultMode`

---

## reset()

---

```php
public function reset($columnKeys, ?int $resultMode = null, 
                      ?bool $styleIdxInclude = null, 
                      ?int $rowLimit = 0): ?Generator
```
_Reset read generator_

### Parameters

* `array|bool|int|null $columnKeys`
* `int|null $resultMode`
* `bool|null $styleIdxInclude`
* `int|null $rowLimit`

---

## rewind()

---

```php
public function rewind($columnKeys, ?int $resultMode = null, 
                       ?bool $styleIdxInclude = null, 
                       ?int $rowLimit = 0): ?Generator
```
_Rewind read generator, alias of reset()_

### Parameters

* `array|bool|int|null $columnKeys`
* `int|null $resultMode`
* `bool|null $styleIdxInclude`
* `int|null $rowLimit`

---

## saveImage()

---

```php
public function saveImage(string $cell, ?string $filename = null): ?string
```
_Writes an image from the cell to the specified filename_

### Parameters

* `string $cell`
* `string|null $filename`

---

## saveImageTo()

---

```php
public function saveImageTo(string $cell, string $dirname): ?string
```
_Writes an image from the cell to the specified directory_

### Parameters

* `string $cell`
* `string $dirname`

---

## setDateFormat()

---

```php
public function setDateFormat($dateFormat): avadim\FastExcelReader\Sheet
```


### Parameters

* `$dateFormat`

---

## setDefaultRowHeight()

---

```php
public function setDefaultRowHeight(float $rowHeight): void
```


### Parameters

* `$rowHeight`

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

### Examples

```php
setReadArea('C3:AZ28'); // set top left and right bottom of read area
setReadArea('C3'); // set top left only
```


---

## setReadAreaColumns()

---

```php
public function setReadAreaColumns(string $columnsRange, 
                                   ?bool $firstRowKeys = false): avadim\FastExcelReader\Sheet
```
_setReadArea('C:AZ') - set left and right columns of read areasetReadArea('C') - set left column only_

### Parameters

* `string $columnsRange`
* `bool|null $firstRowKeys`

---

## setState()

---

```php
public function setState(string $state): avadim\FastExcelReader\Sheet
```


### Parameters

* `string $state`

---

## state()

---

```php
public function state(): string
```


### Parameters

_None_

---

## withHeadings()

---

```php
public function withHeadings(?array $headers = []): SheetReader
```
_Set headings for the sheet_

### Parameters

* `array|null $headers`

---

