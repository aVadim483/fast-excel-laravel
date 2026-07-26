<?php

namespace avadim\FastExcelLaravel;

use avadim\FastExcelReader\AbstractBook;
use avadim\FastExcelReader\AbstractSheet;
use avadim\FastExcelReader\Excel as FastExcelReader;

/**
 * Laravel wrapper around a FastExcelReader book.
 *
 * The wrapper composes an AbstractBook instead of extending a concrete reader, so
 * it works for every format the reader supports (XLSX, XLS, ...): the book is
 * chosen from the file signature and this class relies only on the shared
 * AbstractBook / AbstractSheet public API. Any call the wrapper does not define is
 * delegated to the wrapped book.
 *
 * @mixin AbstractBook
 */
class ExcelReader
{
    protected AbstractBook $book;

    /** @var SheetReader[] Sheet wrappers keyed by the wrapped sheet's object id */
    protected array $sheetWrappers = [];

    /**
     * @param AbstractBook $book
     */
    public function __construct(AbstractBook $book)
    {
        $this->book = $book;
    }

    /**
     * Open a spreadsheet file for import. The reader is selected from the file
     * signature, so XLSX and XLS are opened the same way; the extension is ignored.
     *
     * @param string $file
     * @param array|null $options
     *
     * @return ExcelReader
     */
    public static function open(string $file, ?array $options = []): ExcelReader
    {
        $options = $options ?? [];
        $tempDir = $options['temp_dir'] ?? '';
        if (!$tempDir && function_exists('config')) {
            $tempDir = config('fast-excel.temp_dir') ?: '';
        }

        if (FastExcelReader::isXls($file)) {
            // XLS is read straight from the file, no temp dir is ever used
            $book = FastExcelReader::openXls($file);
        }
        elseif (FastExcelReader::isXlsx($file)) {
            $book = new FastExcelReader($file, $tempDir);
        }
        else {
            // Neither XLS nor XLSX signature: read as CSV. The reader factory
            // returns a CsvBook, and CSV options (delimiter, enclosure, encoding,
            // ...) are passed straight through $options. CSV never uses temp_dir.
            // Global defaults from config('fast-excel.csv') fill the gaps: a
            // per-call option wins, and a null default is skipped so the reader's
            // own default survives.
            if (function_exists('config')) {
                $csvDefaults = config('fast-excel.csv');
                if (is_array($csvDefaults)) {
                    $options += array_filter($csvDefaults, static fn($value) => $value !== null);
                }
            }
            $book = FastExcelReader::open($file, $options);
        }

        return new self($book);
    }

    /**
     * Return the wrapped book instance
     *
     * @return AbstractBook
     */
    public function getBook(): AbstractBook
    {
        return $this->book;
    }

    /**
     * Wrap a sheet returned by the book in a SheetReader, reusing the wrapper so
     * that Laravel-specific state (headings, mapping) set on a sheet survives later
     * access to the same underlying sheet
     *
     * @param AbstractSheet $sheet
     *
     * @return SheetReader
     */
    public function wrapSheet(AbstractSheet $sheet): SheetReader
    {
        $key = spl_object_id($sheet);
        if (!isset($this->sheetWrappers[$key])) {
            $this->sheetWrappers[$key] = new SheetReader($sheet, $this);
        }

        return $this->sheetWrappers[$key];
    }

    /**
     * Get the current or a named sheet as a SheetReader
     *
     * @param string|null $name
     *
     * @return SheetReader|null
     */
    public function sheet(?string $name = null): ?SheetReader
    {
        $sheet = $this->book->sheet($name);

        return $sheet ? $this->wrapSheet($sheet) : null;
    }

    /**
     * Set headings for the current sheet
     *
     * @param array|null $headers
     *
     * @return $this
     */
    public function withHeadings(?array $headers = []): ExcelReader
    {
        $this->sheet()->withHeadings($headers);

        return $this;
    }

    /**
     * Set mapping callback for the current sheet
     *
     * @param $callback
     *
     * @return $this
     */
    public function mapping($callback): ExcelReader
    {
        $this->sheet()->mapping($callback);

        return $this;
    }

    /**
     * Import data into a model from the current sheet
     *
     * @param string $modelClass
     * @param string|bool|null $address
     * @param array|bool|null $columns
     *
     * @return $this
     */
    public function importModel(string $modelClass, $address = null, $columns = null): ExcelReader
    {
        $this->sheet()->importModel($modelClass, $address, $columns);

        return $this;
    }

    /**
     * Read rows from the current sheet (applies the sheet mapping, if any)
     *
     * @param array|bool|int|null $columnKeys
     * @param int|null $resultMode
     * @param bool|null $styleIdxInclude
     *
     * @return array
     */
    public function readRows($columnKeys = [], ?int $resultMode = null, ?bool $styleIdxInclude = null): array
    {
        return $this->sheet()->readRows($columnKeys, $resultMode, $styleIdxInclude);
    }

    /**
     * Delegate every other call to the wrapped book, re-wrapping any sheet(s) it
     * returns so the wrapper stays "sticky" through fluent chains
     *
     * @param string $name
     * @param array $arguments
     *
     * @return mixed
     */
    public function __call(string $name, array $arguments)
    {
        return $this->wrapResult($this->book->$name(...$arguments));
    }

    /**
     * Re-wrap a delegated return value: the book itself becomes $this, a sheet
     * becomes its SheetReader, a list of sheets becomes a list of SheetReaders,
     * anything else passes through unchanged
     *
     * @param mixed $result
     *
     * @return mixed
     */
    protected function wrapResult($result)
    {
        if ($result === $this->book) {
            return $this;
        }
        if ($result instanceof AbstractSheet) {
            return $this->wrapSheet($result);
        }
        if (is_array($result) && $result && !array_filter($result, static fn($item) => !$item instanceof AbstractSheet)) {
            return array_map([$this, 'wrapSheet'], $result);
        }

        return $result;
    }
}
