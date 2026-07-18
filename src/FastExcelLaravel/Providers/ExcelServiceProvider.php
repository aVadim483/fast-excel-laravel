<?php

namespace avadim\FastExcelLaravel\Providers;

use avadim\FastExcelLaravel\Excel;
use Illuminate\Support\ServiceProvider;

class ExcelServiceProvider extends ServiceProvider
{
    /**
     * Bootstrap any application services.
     *
     * @return void
     */
    public function boot()
    {
        if ($this->app->runningInConsole()) {
            $this->publishes([
                __DIR__ . '/../../../config/fast-excel.php' => config_path('fast-excel.php'),
            ], 'config');
        }
    }

    /**
     * Register any application services.
     *
     * @return void
     */
    public function register()
    {
        $this->mergeConfigFrom(__DIR__ . '/../../../config/fast-excel.php', 'fast-excel');

        $this->app->bind(Excel::class);
        $this->app->alias(Excel::class, 'excel');
    }
}
