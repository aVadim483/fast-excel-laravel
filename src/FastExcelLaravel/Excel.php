<?php

namespace avadim\FastExcelLaravel;

class Excel
{
    /**
     * @param array|null $options
     */
    public function __construct(?array $options = [])
    {

    }

    /**
     * @param $sheets
     * @param array|null $options
     *
     * @return ExcelWriter
     */
    public static function create($sheets = null, ?array $options = []): ExcelWriter
    {
        return ExcelWriter::create($sheets, $options);
    }

    /**
     * @param $sheets
     * @param array|null $options
     *
     * @return ExcelReader
     */
    public static function open($sheets = null, ?array $options = []): ExcelReader
    {
        return ExcelReader::open($sheets, $options);
    }
}
