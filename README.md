# FastExcelLaravel
Lightweight and very fast XLSX Excel Spreadsheet Writer for Laravel 
(wrapper for [FastExcelWriter](https://packagist.org/packages/avadim/fast-excel-writer))

## Introduction

Exporting data from your Laravel application has never been so fast! Importing models into your Laravel application has never been so easy!

This library is a wrapper for avadim/fast-excel-writer фтв avadim/fast-excel-куфвук, so it's also lightweight, fast, and requires a minimum of memory.
Using this library, you can export arrays, collections and models to a XLSX-file from your Laravel application.

## Installation

Install via composer:

```
composer require avadim/fast-excel-laravel
```
And then you can use facade ```Excel```

```php
// Create workbook
$excel = \Excel::create();
// do something...
// and save XLSX-file to default storage
$excel->saveTo('path/file.xlsx');

// or save file to specified disk
$excel->store('disk', 'path/file.xlsx');
```
Export a Model
```php

// Create workbook with sheet named 'Users'
$excel = \Excel::create('Users');

// Write all users to Excel file
$sheet->writeModel(Users::class);

// Write users with field names in the first row
$sheet->withHeaders()
    ->applyFontStyleBold()
    ->applyBorder('thin')
    ->writeModel(Users::class);

$excel->saveTo('path/file.xlsx');
```

Export any collections and arrays
```php
// Create workbook with sheet named 'Users'
$excel = \Excel::create('Users');

$sheet = $excel->getSheet();
$users = User::all();
$sheet->writeData($users);

$sheet = $excel->makeSheet('Records');
$records = \DB::table('users')->where('age', '>=', 21)->get(['id', 'name', 'birthday']);
$sheet->writeData($records);

$sheet = $excel->makeSheet('Collection');
$collection = collect([
    [ 'id' => 1, 'site' => 'google.com' ],
    [ 'id' => 2, 'site.com' => 'youtube.com' ],
]);
$sheet->writeData($collection);

$sheet = $excel->makeSheet('Array');
$array = [
    [ 'id' => 1, 'name' => 'Helen' ],
    [ 'id' => 2, 'name' => 'Peter' ],
];
$sheet->writeData($array);

$sheet = $excel->makeSheet('Callback');
$sheet->writeData(function () {
    foreach (User::cursor() as $user) {
        yield $user;
    }
});

```

## Advanced usage for data export

See detailed documentation for avadim/fast-excel-laravel here: https://github.com/aVadim483/fast-excel-writer/tree/master#readme

```php
$excel = \Excel::create('Users');
$sheet = $excel->getSheet();

$sheet->setColWidth('B', 12);
$sheet->setColOptions('c', ['width' => 12, 'text-align' => 'center']);
$sheet->setColWidth('d', 'auto');

$title = 'This is demo of avadim/fast-excel-laravel';
// begin area for direct access to cells
$area = $sheet->beginArea();
$area->setValue('A2:D2', $title)
      ->applyFontSize(14)
      ->applyFontStyleBold()
      ->applyTextCenter();

$area
    ->setValue('a4:a5', '#')
    ->setValue('b4:b5', 'Number')
    ->setValue('c4:d4', 'Movie Character')
    ->setValue('c5', 'Birthday')
    ->setValue('d5', 'Name')
;

$area->withRange('a4:d5')
    ->applyBgColor('#ccc')
    ->applyFontStyleBold()
    ->applyOuterBorder('thin')
    ->applyInnerBorder('thick')
    ->applyTextCenter();

$sheet->writeAreas();

$sheet->writeData($data);
$excel->saveTo($testFileName);

```

## Do you want to support FastExcelLaravel?

if you find this package useful you can support and donate to me for a cup of coffee:

* USDT (TRC20) TSsUFvJehQBJCKeYgNNR1cpswY6JZnbZK7
* USDT (ERC20) 0x5244519D65035aF868a010C2f68a086F473FC82b

Or just give me star on GitHub :)
