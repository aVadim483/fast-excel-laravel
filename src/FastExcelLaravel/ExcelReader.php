<?php

namespace avadim\FastExcelLaravel;

class ExcelReader extends \avadim\FastExcelReader\Excel
{
    /**
     * Open XLSX file for import
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
     * Create SheetReader instance
     *
     * @param string $sheetName
     * @param $sheetId
     * @param $file
     * @param $path
     *
     * @return SheetReader
     */
    public static function createSheet(string $sheetName, $sheetId, $file, $path, $excel): SheetReader
    {
        return new SheetReader($sheetName, $sheetId, $file, $path, $excel);
    }

    /**
     * Set headings for the current sheet
     *
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
     * Set mapping callback for the current sheet
     *
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
     * Import data into a model from the current sheet
     *
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
