<?php

namespace avadim\FastExcelLaravel;

class SheetReader extends \avadim\FastExcelReader\Sheet
{
    public function loadModel($modelClass, $address = null, $columns = null)
    {
        $this->setReadArea($address);
        foreach ($this->nextRow() as $rowNum => $rowData) {
            $model = new $modelClass;
        }
    }
}