<?php

namespace avadim\FastExcelLaravel;

use avadim\FastExcelReader\AbstractSheet;
use avadim\FastExcelReader\Excel as FastExcelReader;
use Illuminate\Database\Eloquent\Model;

/**
 * Laravel wrapper around a FastExcelReader sheet.
 *
 * Composes an AbstractSheet (XLSX, XLS, ...) and adds the Eloquent-aware import
 * helpers. Every call the wrapper does not define is delegated to the wrapped
 * sheet; a sheet returned by such a call is re-wrapped through the owning book so
 * fluent chains keep their Laravel-specific state.
 *
 * @mixin AbstractSheet
 */
class SheetReader
{
    protected AbstractSheet $sheet;

    protected ExcelReader $excel;

    private int $resultMode = 0;

    private array $customHeaders = [];

    /** @var mixed|null */
    private $mappingCallback = null;

    /**
     * @param AbstractSheet $sheet
     * @param ExcelReader $excel
     */
    public function __construct(AbstractSheet $sheet, ExcelReader $excel)
    {
        $this->sheet = $sheet;
        $this->excel = $excel;
    }

    /**
     * Return the wrapped sheet instance
     *
     * @return AbstractSheet
     */
    public function getSheet(): AbstractSheet
    {
        return $this->sheet;
    }

    /**
     * Set headings for the sheet. The first row of the read area is always skipped;
     * if $headers is not empty, these names are used as attribute names (in column order)
     * instead of the values of the first row
     *
     * Named withHeadings() to match the Laravel ecosystem convention; the underlying
     * reader calls the same concept withHeader() (AbstractSheet::withHeader()).
     *
     * @param array|null $headers
     *
     * @return $this
     */
    public function withHeadings(?array $headers = []): SheetReader
    {
        $this->resultMode = FastExcelReader::KEYS_FIRST_ROW;
        $this->customHeaders = $headers ? array_values($headers) : [];

        return $this;
    }

    /**
     * Set mapping callback for the sheet
     *
     * @param $callback
     *
     * @return $this
     */
    public function mapping($callback): SheetReader
    {
        if (is_array($callback)) {
            $mapArray = $callback;
            $callback = function ($row) use($mapArray) {
                $record = [];
                foreach ($row as $col => $value) {
                    if (isset($mapArray[$col])) {
                        $record[$mapArray[$col]] = $value;
                    }
                    else {
                        $record[$col] = $value;
                    }
                }
                return $record;
            };
        }
        $this->mappingCallback = $callback;

        return $this;
    }

    /**
     * Load models from Excel to database
     *      loadModels(User::class)
     *      loadModels(User::class, true) -- the first row used as a field names
     *      loadModels(User::class, 'B:D') -- read data from columns B:D
     *      loadModels(User::class, 'B3') -- read data from area started at B3
     *      loadModels(User::class, 'B3', true) -- read data from area started at B3 and the first row used as a field names
     *
     * @param $modelClass
     * @param $address
     * @param $columns
     *
     * @return $this
     */
    public function importModel($modelClass, $address = null, $columns = null): SheetReader
    {
        if ($address && is_string($address)) {
            $this->sheet->setReadArea($address);
        }
        foreach ($this->sheet->nextRow($columns, $this->resultMode) as $rowData) {
            /** @var Model $model */
            $model = new $modelClass;
            if ($this->customHeaders) {
                $rowData = $this->_applyCustomHeaders($rowData);
            }
            if ($this->mappingCallback) {
                $rowData = call_user_func($this->mappingCallback, $rowData);
            }
            $model->fill($rowData);
            $model->save();
        }
        $this->resultMode = 0;
        $this->customHeaders = [];

        return $this;
    }

    /**
     * Replace row keys with custom header names (in column order)
     *
     * @param array $rowData
     *
     * @return array
     */
    protected function _applyCustomHeaders(array $rowData): array
    {
        $values = array_values($rowData);
        $record = [];
        foreach ($this->customHeaders as $idx => $attribute) {
            $record[$attribute] = $values[$idx] ?? null;
        }

        return $record;
    }

    /**
     * Returns cell values as a two-dimensional array (applies the mapping, if any)
     *
     * @param array|bool|int|null $columnKeys
     * @param int|null $resultMode
     * @param bool|null $styleIdxInclude
     *
     * @return array
     */
    public function readRows($columnKeys = [], ?int $resultMode = null, ?bool $styleIdxInclude = null): array
    {
        $rows = $this->sheet->readRows($columnKeys, $resultMode, $styleIdxInclude);
        if ($this->mappingCallback) {
            foreach ($rows as $rowNum => $rowData) {
                $rows[$rowNum] = call_user_func($this->mappingCallback, $rowData);
            }
        }

        return $rows;
    }

    /**
     * Delegate every other call to the wrapped sheet, re-wrapping any sheet it
     * returns so fluent chains keep the wrapper (and its state)
     *
     * @param string $name
     * @param array $arguments
     *
     * @return mixed
     */
    public function __call(string $name, array $arguments)
    {
        $result = $this->sheet->$name(...$arguments);
        if ($result === $this->sheet) {
            return $this;
        }
        if ($result instanceof AbstractSheet) {
            return $this->excel->wrapSheet($result);
        }

        return $result;
    }
}
