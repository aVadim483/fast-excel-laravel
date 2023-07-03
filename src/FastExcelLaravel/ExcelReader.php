<?php

namespace avadim\FastExcelLaravel;

class ExcelReader extends \avadim\FastExcelReader\Excel
{
    public static function createSheet(string $sheetName, $sheetId, $file, $path): SheetReader
    {
        return new SheetReader($sheetName, $sheetId, $file, $path);
    }

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

    /**
     * @param array|null $headers
     *
     * @return $this
     */
    public function withHeadings(?array $headers = []): ExcelReader
    {
        $this->sheet()->withHeadings($headers);

        return $this;
    }

    /**
     * @param $callback
     *
     * @return $this
     */
    public function mapping($callback): ExcelReader
    {
        $this->sheet()->mapping($callback);

        return $this;
    }

    /**
     * @param string $modelClass
     * @param string|bool|null $address
     * @param array|bool|null $columns
     *
     * @return $this
     */
    public function importModel(string $modelClass, $address = null, $columns = null): ExcelReader
    {
        $this->sheet()->importModel($modelClass, $address, $columns);

        return $this;
    }
}
