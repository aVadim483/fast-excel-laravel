<?php

namespace avadim\FastExcelLaravel;

use Illuminate\Database\Eloquent\Model;

class SheetReader extends \avadim\FastExcelReader\Sheet
{
    private int $resultMode = 0;
    private $mappingCallback = null;

    /**
     * @param array|null $headers
     *
     * @return $this
     */
    public function withHeadings(?array $headers = []): SheetReader
    {
        $this->resultMode = \avadim\FastExcelReader\Excel::KEYS_FIRST_ROW;

        return $this;
    }

    /**
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
     * Load models from Excel
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
            if ($this->mappingCallback) {
                $rowData = call_user_func($this->mappingCallback, $rowData);
            }
            $model->fill($rowData);
            $model->save();
        }
        $this->resultMode = 0;

        return $this;
    }
}