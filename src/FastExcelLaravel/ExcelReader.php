<?php

namespace avadim\FastExcelLaravel;

class ExcelReader extends \avadim\FastExcelReader\Excel
{
    /**
     * Open XLSX file
     *
     * @param string $file
     *
     * @return ExcelReader
     */
    public static function open(string $file): ExcelReader
    {
        return new self($file);
    }

}
