<table border="0">
<tr>
<td valign="top"><img height="240px" src="logo/logo-laravel.jpg" alt="FastExcelLaravel"></td>
<td valign="top">
<p align="center">

# FastExcelLaravel

</p>
</td>
</tr>
</table>

[English](README.md) | **Русский**

[![Tests](https://github.com/aVadim483/fast-excel-laravel/actions/workflows/tests.yml/badge.svg)](https://github.com/aVadim483/fast-excel-laravel/actions/workflows/tests.yml)
[![Latest Stable Version](https://img.shields.io/packagist/v/avadim/fast-excel-laravel)](https://packagist.org/packages/avadim/fast-excel-laravel)
[![Total Downloads](https://img.shields.io/packagist/dt/avadim/fast-excel-laravel)](https://packagist.org/packages/avadim/fast-excel-laravel)
[![License](https://img.shields.io/packagist/l/avadim/fast-excel-laravel)](https://packagist.org/packages/avadim/fast-excel-laravel)

Лёгкая и очень быстрая библиотека чтения/записи XLSX-файлов Excel для Laravel на чистом PHP
(обёртка над [FastExcelWriter](https://packagist.org/packages/avadim/fast-excel-writer)
 и [FastExcelReader](https://packagist.org/packages/avadim/fast-excel-reader))

## Введение

Экспорт данных из вашего Laravel-приложения ещё никогда не был таким быстрым! Импорт моделей в ваше Laravel-приложение ещё никогда не был таким простым!

Эта библиотека — обёртка над **avadim/fast-excel-writer** и **avadim/fast-excel-reader**, поэтому она такая же лёгкая, быстрая и требует минимум памяти.
С её помощью вы можете экспортировать массивы, коллекции и модели в XLSX-файл из вашего Laravel-приложения, а также импортировать данные обратно.

**Возможности**

* Запись
  * Простой экспорт моделей, коллекций и массивов в Excel
  * Очень быстрый экспорт огромных наборов данных с минимальным потреблением памяти
  * Создание нескольких листов, базовые стили колонок, строк и ячеек
  * Настройка высоты строк и ширины колонок (включая автоматический расчёт ширины)
  * Маппинг экспортируемых данных
  * Активные гиперссылки, формулы, заметки и изображения в результирующих XLSX-файлах
  * Защита книги и листов с паролем и без
  * Параметры страницы — поля, размер страницы
  * Вставка диаграмм
  * Валидация данных и условное форматирование
* Чтение
  * Очень быстрый импорт книг и листов в Eloquent-модели с минимальным потреблением памяти
  * Автоматическое определение полей по заголовкам импортируемой таблицы
  * Импорт очень больших файлов с минимальным потреблением памяти
  * Маппинг импортируемых данных
  * Автоматическое и настраиваемое форматирование значений даты и времени при импорте
  * Извлечение изображений из XLSX-файлов
* Маппинг данных при импорте и экспорте

## Требования

| Версия | PHP    | Laravel |
|--------|--------|---------|
| 3.x    | >= 8.1 | 10 – 13 |
| 2.x    | >= 7.4 | 6 – 13  |

## Установка

Установите через composer:

```
composer require avadim/fast-excel-laravel
```
После этого вы можете использовать фасад ```Excel```

```php
// Create a new workbook...
$excel = \Excel::create();

// export model...
$excel->sheet()->withHeadings()->exportModel(Users::class);

// and save XLSX-file to default storage
$excel->saveTo('path/file.xlsx');

// or save file to specified disk
$excel->store('disk', 'path/file.xlsx');

// Open an existing workbook
$excel = \Excel::open(storage_path('path/file.xlsx'));

// import records to database
$excel->withHeadings()->importModel(User::class);
```

Содержание:
* [Экспорт данных](#экспорт-данных)
  * [Экспорт модели](#экспорт-модели)
  * [Экспорт коллекций и массивов](#экспорт-коллекций-и-массивов)
  * [Маппинг экспортируемых данных](#маппинг-экспортируемых-данных)
  * [Расширенные возможности экспорта](#расширенные-возможности-экспорта)
  * [Отдача файла на скачивание](#отдача-файла-на-скачивание)
* [Импорт данных](#импорт-данных)
  * [Импорт модели](#импорт-модели)
  * [Маппинг импортируемых данных](#маппинг-импортируемых-данных)
  * [Расширенные возможности импорта](#расширенные-возможности-импорта)
* [Конфигурация](#конфигурация)
* [Дополнительные возможности](#дополнительные-возможности)
* [Хотите поддержать FastExcelLaravel?](#хотите-поддержать-fastexcellaravel)

[Справочник API находится здесь](/docs/90-api-reference.md)

## Экспорт данных

### Экспорт модели
Простой и быстрый экспорт модели. В этом случае экспортируются только данные модели — без заголовков и какого-либо оформления
```php

// Create workbook with sheet named 'Users'
$excel = \Excel::create('Users');
$sheet = $excel->sheet();

// Export all users to Excel file
$sheet->exportModel(Users::class);

$excel->saveTo('path/file.xlsx');
```
Следующий код запишет в первую строку имена полей со стилями (шрифт и границы), а затем экспортирует все данные модели User

```php

// Create workbook with sheet named 'Users'
$excel = \Excel::create('Users');
$sheet = $excel->sheet();

// Write users with field names in the first row
$sheet->withHeadings()
    ->applyFontStyleBold()
    ->applyBorder('thin')
    ->exportModel(Users::class);

$excel->saveTo('path/file.xlsx');
```

### Экспорт коллекций и массивов
```php
// Create workbook with sheet named 'Users'
$excel = \Excel::create('Users');

$sheet = $excel->sheet();
// Get users as collection
$users = User::where('age', '>', 35)->get();

// Write attribute names
$sheet->writeRow(array_keys($users->first()->getAttributes()));

// Write all selected records
$sheet->writeData($users);

$sheet = $excel->makeSheet('Records');
// Get collection of records using Query Builder
$records = \DB::table('users')->where('age', '>=', 21)->get(['id', 'name', 'birthday']);
$sheet->writeData($records);

$sheet = $excel->makeSheet('Collection');
// Make custom collection of arrays
$collection = collect([
    [ 'id' => 1, 'site' => 'google.com' ],
    [ 'id' => 2, 'site.com' => 'youtube.com' ],
]);
$sheet->writeData($collection);

$sheet = $excel->makeSheet('Array');
// Make array and write to sheet
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

### Маппинг экспортируемых данных

Вы можете преобразовать данные, которые будут добавлены в строку

```php
$sheet = $excel->sheet();
$sheet->mapping(function($model) {
    return [
        'id' => $model->id, 'date' => $model->created_at, 'name' => $model->first_name . $model->last_name,
    ];
})->exportModel(User::class);
$excel->save($testFileName);

```

### Расширенные возможности экспорта

Подробная документация по avadim/fast-excel-writer находится здесь: https://github.com/aVadim483/fast-excel-writer/tree/master#readme

```php
$excel = \Excel::create('Users');
$sheet = $excel->sheet();

// Set column B to 12
$sheet->setColWidth('B', 12);
// Set styles for column C
$sheet->setColDataStyle('C', ['width' => 12, 'text-align' => 'center']);
// Set column width to auto
$sheet->setColWidth('D', 'auto');

$title = 'This is demo of avadim/fast-excel-laravel';
// Begin area for direct access to cells
$area = $sheet->beginArea();
$area->setValue('A2:D2', $title)
      ->applyFontSize(14)
      ->applyFontStyleBold()
      ->applyTextCenter();
      
// Write headers to area, column letters are case independent
$area
    ->setValue('a4:a5', '#')
    ->setValue('b4:b5', 'Number')
    ->setValue('c4:d4', 'Movie Character')
    ->setValue('c5', 'Birthday')
    ->setValue('d5', 'Name')
;

// Apply styles to headers
$area->withRange('a4:d5')
    ->applyBgColor('#ccc')
    ->applyFontStyleBold()
    ->applyOuterBorder('thin')
    ->applyInnerBorder('thick')
    ->applyTextCenter();
    
// Write area to sheet
$sheet->writeAreas();

// You can set value formats for some fields
$sheet->formatAttributes(['birthday' => '@date', 'number' => '@integer']);

// Write data to sheet
$sheet->writeData($data);

// Save XLSX-file
$excel->saveTo($testFileName);

```

### Отдача файла на скачивание

Вы можете вернуть сгенерированный файл как download-ответ прямо из экшена контроллера

```php
public function export()
{
    $excel = \Excel::create('Users');
    $excel->sheet()->withHeadings()->exportModel(User::class);

    // Returns Symfony\Component\HttpFoundation\BinaryFileResponse,
    // the temporary file will be deleted after sending
    return $excel->download('users.xlsx');
}
```

## Импорт данных

### Импорт модели
Для импорта моделей используйте метод ```importModel()```.
Если первая строка содержит имена полей, примените их с помощью метода ```withHeadings()```

![import.jpg](import.jpg)

```php
// Open XLSX-file 
$excel = Excel::open($file);

// Import a workbook to User model using the first row as attribute names
$excel->withHeadings()->importModel(User::class);

// Done!!!
```
Вы также можете задать собственные имена атрибутов (в порядке колонок) — первая строка по-прежнему пропускается,
но её значения игнорируются
```php
$excel->withHeadings(['name', 'birthday', 'random'])->importModel(User::class);
```
Вы можете указать колонки или ячейки, из которых будет выполняться импорт

```php
// Import row to User model from columns range A:B - only 'name' and 'birthday'
$excel->withHeadings()->importModel(User::class, 'A:B');
```

![import2.jpg](import2.jpg)
```php
// Import from cells range
$excel->withHeadings()->importModel(User::class, 'B4:D7');

// Define top left cell only
$excel->withHeadings()->importModel(User::class, 'B4');
```
В последних двух примерах также предполагается, что первая строка импортируемых данных (строка 4) —
это имена атрибутов.

### Маппинг импортируемых данных

Вы можете сами задать соответствие между колонками и именами полей.

```php
// Import row to User model from columns range B:E
$excel->mapping(function ($record) {
    return [
        'id' => $record['A'], 'name' => $record['B'], 'birthday' => $record['C'], 'random' => $record['D'],
    ];
})->importModel(User::class, 'B:D');

// Define top left cell only
$excel->mapping(['B' => 'name', 'C' => 'birthday', 'D' => 'random'])->importModel(User::class, 'B5');

// Define top left cell only (shorter way)
$excel->importModel(User::class, 'B5', ['B' => 'name', 'C' => 'birthday', 'D' => 'random']);
```

### Расширенные возможности импорта
Подробная документация по avadim/fast-excel-reader находится здесь: https://github.com/aVadim483/fast-excel-reader/tree/master#readme
```php
$excel = Excel::open($file);

$sheet = $excel->sheet('Articles');
$sheet->setReadArea('B5');
foreach ($sheet->nextRow() as $rowNum => $rowData) {
    $user = User::create([
        'name' => $rowData['B'],
        'birthday' => new \Carbon\Carbon($rowData['C']),
        'password' => bcrypt($rowData['D']),
    ]);
    Article::create([
        'user_id' => $user->id,
        'title' => $rowData['E'],
        'date' => new \Carbon\Carbon($rowData['F']),
        'public' => $rowData['G'] === 'yes',
    ]);
}
```

## Конфигурация

При необходимости вы можете опубликовать конфигурационный файл

```
php artisan vendor:publish --provider="avadim\FastExcelLaravel\Providers\ExcelServiceProvider" --tag=config
```

Доступные опции в `config/fast-excel.php`:

```php
return [
    // Directory for temporary files created while reading and writing XLSX.
    // When null, storage_path('app/tmp/fast-excel') is used
    'temp_dir' => null,
];
```

Устаревшие временные файлы (старше 24 часов, оставшиеся после сбоев) удаляются автоматически.

Также можно переопределить каталог временных файлов для отдельной книги через опции

```php
$excel = \Excel::create('Users', ['temp_dir' => '/path/to/tmp']);
$excel = \Excel::open($file, ['temp_dir' => '/path/to/tmp']);
```

## Дополнительные возможности
Дополнительные возможности экспорта описаны в документации [FastExcelWriter](https://packagist.org/packages/avadim/fast-excel-writer).

Дополнительные возможности импорта описаны в документации [FastExcelReader](https://packagist.org/packages/avadim/fast-excel-reader)

## Хотите поддержать FastExcelLaravel?

Если этот пакет оказался полезным, вы можете поддержать меня и задонатить на чашку кофе:

* USDT (TRC20) TSsUFvJehQBJCKeYgNNR1cpswY6JZnbZK7
* USDT (ERC20) 0x5244519D65035aF868a010C2f68a086F473FC82b
* ETH 0x5244519D65035aF868a010C2f68a086F473FC82b 

Или просто поставьте звёздочку на GitHub :)
