# Изменения

[🇬🇧 English](CHANGELOG.md) · 🇷🇺 Русский

<!-- NOTE for maintainers: this file exists in two languages.
     When editing CHANGELOG.ru.md, please update CHANGELOG.md accordingly.
     Code and identifiers must stay byte-for-byte identical in both versions; only prose is translated. -->

Все значимые изменения проекта фиксируются в этом файле.

Формат основан на [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
проект придерживается [семантического версионирования](https://semver.org/spec/v2.0.0.html).
Более раннюю историю см. на
[странице релизов](https://github.com/aVadim483/fast-excel-laravel/releases).

## 4.2.0 — 2026-08-16

### Добавлено

* **Чтение книги из строки или потока.** `\Excel::openString($content)` открывает книгу, лежащую в строке
  (поле в базе, тело HTTP-ответа, `Storage::get()`, …), а `\Excel::openStream($stream)` — книгу за любым
  открытым потоком (`Storage::readStream()`, `fopen('https://…')`, `php://memory`, …). Содержимое копируется
  во временный файл и открывается точно так же, как в `\Excel::open()`, поэтому формат по-прежнему
  определяется по содержимому (XLSX, XLS, CSV), действуют те же опции, а дальше работает весь API импорта.
  Временная копия удаляется по завершении скрипта; поток не закрывается. См.
  [docs/82-reading-from-memory.md](docs/82-reading-from-memory.md).

### Изменено

* Требование к `avadim/fast-excel-writer` поднято до `^6.16` (исправления предрасчитанных результатов формул,
  утечки памяти в долгоживущих процессах, значений длиннее 32767 символов и сломанного XML из-за
  неэкранированных символов в правилах валидации и условном форматировании).
* Требование к `avadim/fast-excel-reader` поднято до `^4.4.1`. Writer экранирует управляющие символы как
  `_xHHHH_`, и только эта версия ридера декодирует их обратно, поэтому цикл «экспорт → импорт» сохраняет такие
  значения (ячейка с CR раньше читалась как буквальный текст `_x000D_`).

## 4.1.0 — 2026-07-26

### Добавлено

* **Чтение CSV-файлов.** `\Excel::open()` теперь читает CSV в дополнение к XLSX и устаревшему XLS. Файл,
  который не является ни корректным XLSX (ZIP), ни XLS (OLE2), читается как CSV, поэтому весь API импорта
  (`importModel()`, `withHeadings()`, `mapping()`, области чтения, `readRows()`, `nextRow()`, …) работает
  одинаково. Значения CSV читаются как строки, стилей у CSV нет; запись CSV не поддерживается (экспорт
  по-прежнему только в XLSX). См. [docs/81-reading-csv.md](docs/81-reading-csv.md).
* **CSV-опции.** `\Excel::open()` принимает CSV-опции (`delimiter`, `enclosure`, `escape`, `encoding`,
  `skip_empty_lines`, `comment_prefix`, `mode`) вторым аргументом, с глобальными значениями по умолчанию в
  новой секции `csv` файла `config/fast-excel.php`.

### Изменено

* Требование к `avadim/fast-excel-reader` поднято до `^4.3` (добавляет CSV-ридер).

## 4.0.0 — 2026-07-25

### Добавлено

* **Чтение устаревших книг XLS (Excel 97-2003, BIFF8).** `\Excel::open()` теперь принимает и XLS, и XLSX;
  формат определяется по сигнатуре файла, поэтому расширение игнорируется, а весь API импорта
  (`importModel()`, `withHeadings()`, `mapping()`, области чтения, `readRows()`, `nextRow()`, …) работает для
  обоих форматов одинаково. Запись XLS не поддерживается — экспорт по-прежнему только в XLSX. См.
  [docs/80-reading-xls.md](docs/80-reading-xls.md).

### Изменено

* Требование к `avadim/fast-excel-reader` поднято до `^4.0`.
* **`ExcelReader` и `SheetReader` теперь используют композицию вместо наследования.** Они оборачивают
  `avadim\FastExcelReader\AbstractBook` / `AbstractSheet` и делегируют любой не определённый в обёртке вызов,
  поэтому одна реализация обслуживает все форматы reader'а. Классы writer'а не изменились.

### Несовместимые изменения (BREAKING)

* **`ExcelReader` больше не наследует `\avadim\FastExcelReader\Excel`, а `SheetReader` больше не наследует
  `\avadim\FastExcelReader\Sheet`.** Поэтому объекты, возвращаемые `\Excel::open()` и `->sheet()`, больше не
  являются экземплярами этих классов reader'а.
* Статические хелперы, ранее унаследованные обёрткой (`ExcelReader::validate()`, `::colLetter()`,
  `::colNum()`, `::createReader()`, `::setTempDir()`, …), удалены. Вызывайте их напрямую у
  `\avadim\FastExcelReader\Excel`.
* Константы класса, ранее унаследованные обёрткой (`ExcelReader::KEYS_*`), удалены. Обращайтесь к ним через
  `\avadim\FastExcelReader\Excel` (этот путь работал всегда и не изменился).
* Внутренняя фабрика `ExcelReader::createSheet()` удалена (reader больше не создаёт листы обёртки).

### Руководство по обновлению (3.x → 4.x)

* Обычное использование не меняется: `\Excel::open()`, `->withHeadings()`, `->mapping()`, `->importModel()`,
  `->readRows()`, `->from()`, `->setDateFormat()`, `['temp_dir' => …]` и остальные документированные вызовы
  работают как прежде.
* Если ваш код тайпхинтит или проверяет `instanceof` на `\avadim\FastExcelReader\Excel` / `\Sheet` для
  значений из этого пакета, замените их на `\avadim\FastExcelLaravel\ExcelReader` /
  `\avadim\FastExcelLaravel\SheetReader`.
* Если ваш код обращался к унаследованным статическим хелперам или константам `KEYS_*` через класс обёртки,
  перенесите эти обращения на `\avadim\FastExcelReader\Excel`.
