<?php

spl_autoload_register(static function ($class) {
    $namespace = 'avadim\\FastExcelLaravel\\';
    if (0 === strpos($class, $namespace)) {
        include __DIR__ . '/FastExcelLaravel/' . str_replace($namespace, '', $class) . '.php';
    }
});

// EOF