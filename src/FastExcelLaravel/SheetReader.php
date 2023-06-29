<?php

namespace avadim\FastExcelLaravel;

use Illuminate\Database\Eloquent\Model;

class SheetReader extends \avadim\FastExcelReader\Sheet
{
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
    public function loadModels($modelClass, $address = null, $columns = null): SheetReader
    {
        $resultMode = 0;
        if ($columns === true) {
            $resultMode = \avadim\FastExcelReader\Excel::KEYS_FIRST_ROW;
            $columns = [];
        }
        elseif ($columns === null && $address === true) {
            $columns = [];
            $resultMode = \avadim\FastExcelReader\Excel::KEYS_FIRST_ROW;
        }
        if ($address && is_string($address)) {
            $this->setReadArea($address);
        }
        $tz = date_default_timezone_get();
        foreach ($this->nextRow($columns, $resultMode) as $rowData) {
            /** @var Model $model */
            $model = new $modelClass;
            $model->fill($rowData);
            $model->save();
        }

        return $this;
    }
}