<?php

namespace avadim\FastExcelLaravel;

use Illuminate\Database\Eloquent\Model;

class SheetReader extends \avadim\FastExcelReader\Sheet
{
    private int $resultMode = 0;

    private array $customHeaders = [];

    /** @var mixed|null */
    private $mappingCallback = null;

    /**
     * Set headings for the sheet. The first row of the read area is always skipped;
     * if $headers is not empty, these names are used as attribute names (in column order)
     * instead of the values of the first row
     *
     * @param array|null $headers
     *
     * @return $this
     */
    public function withHeadings(?array $headers = []): SheetReader
    {
        $this->resultMode = \avadim\FastExcelReader\Excel::KEYS_FIRST_ROW;
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
            $this->setReadArea($address);
        }
        foreach ($this->nextRow($columns, $this->resultMode) as $rowData) {
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
     * Returns cell values as a two-dimensional array
     *
     * @param array|bool|int|null $columnKeys
     * @param int|null $resultMode
     * @param bool|null $styleIdxInclude
     *
     * @return array
     */
    public function readRows($columnKeys = [], ?int $resultMode = null, ?bool $styleIdxInclude = null): array
    {
        $rows = parent::readRows($columnKeys, $resultMode, $styleIdxInclude);
        if ($this->mappingCallback) {
            foreach ($rows as $rowNum => $rowData) {
                $rows[$rowNum] = call_user_func($this->mappingCallback, $rowData);
            }
        }

        return $rows;
    }
}