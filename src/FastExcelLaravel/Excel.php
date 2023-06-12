<?php

namespace avadim\FastExcelLaravel;

class Excel
{
    public function __construct(?array $options = [])
    {

    }

    public static function create($sheets = null, ?array $options = [])
    {
        return ExcelWriter::create($sheets);
    }
}
