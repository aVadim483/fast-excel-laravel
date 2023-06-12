<?php

namespace avadim\FastExcelLaravel\Providers;

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
        //
    }

    /**
     * Register any application services.
     *
     * @SuppressWarnings("unused")
     *
     * @return void
     */
    public function register()
    {
        $this->app->bind('excel', function ($app, $data = null) {
            if (is_array($data)) {
                $data = collect($data);
            }

            return new \avadim\FastExcelLaravel\Excel();
        });
    }
}
