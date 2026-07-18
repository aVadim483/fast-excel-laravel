<?php

namespace avadim\FastExcelLaravel\Facades;

use Illuminate\Support\Facades\Facade;

/**
 * Class Excel
 *
 * @method static \avadim\FastExcelLaravel\ExcelWriter create($sheets = null, ?array $options = [])
 * @method static \avadim\FastExcelLaravel\ExcelReader open(string $file, ?array $options = [])
 *
 * @see \avadim\FastExcelLaravel\Excel
 */
class Excel extends Facade
{
    /**
     * Get the registered name of the component.
     *
     * @return string
     */
    protected static function getFacadeAccessor()
    {
        return 'excel';
    }
}
