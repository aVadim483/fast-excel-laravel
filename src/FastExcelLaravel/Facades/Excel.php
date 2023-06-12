<?php

namespace avadim\FastExcelLaravel\Facades;

use Illuminate\Support\Facades\Facade;

/**
 * Class Excel
 *
 * @method static string export($path, callable $callback = null)
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
