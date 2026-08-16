<?php

namespace avadim\FastExcelLaravel;

use avadim\FastExcelReader\AbstractBook;
use avadim\FastExcelReader\AbstractSheet;
use avadim\FastExcelReader\Excel as FastExcelReader;
use avadim\FastExcelReader\Exception;

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
     * Open a spreadsheet held in a string, e.g. a database blob, an HTTP response
     * body or a file read from a Laravel disk with Storage::get()
     *
     * The content is written to a temporary file and then opened by open(), so the
     * format is detected from the bytes (XLSX, XLS or CSV) and every option open()
     * accepts works here too. The temporary file is removed on script shutdown.
     *
     * @param string $content Raw bytes of the workbook
     * @param array|null $options Same options as open()
     *
     * @return ExcelReader
     */
    public static function openString(string $content, ?array $options = []): ExcelReader
    {
        if ($content === '') {
            throw new Exception('Cannot open an empty string as a spreadsheet');
        }
        $tempFile = self::makeTempFile($options);
        if (file_put_contents($tempFile, $content) === false) {
            @unlink($tempFile);
            throw new Exception('Cannot write the spreadsheet content to temporary file "' . $tempFile . '"');
        }

        return self::openTempFile($tempFile, $options);
    }

    /**
     * Open a spreadsheet from an open stream resource, e.g. a remote file or a
     * Laravel disk read: \Excel::openStream(Storage::disk('s3')->readStream($path))
     *
     * The stream is copied into a temporary file from its current position (without
     * seeking, so non-rewindable streams such as HTTP wrappers work) and then opened
     * by open(). The caller keeps ownership of the stream, it is not closed here;
     * the temporary file is removed on script shutdown.
     *
     * @param resource $stream An open readable stream resource
     * @param array|null $options Same options as open()
     *
     * @return ExcelReader
     */
    public static function openStream($stream, ?array $options = []): ExcelReader
    {
        if (!is_resource($stream)) {
            throw new Exception('openStream() expects an open stream resource');
        }
        $tempFile = self::makeTempFile($options);
        $out = fopen($tempFile, 'wb');
        if (!$out) {
            @unlink($tempFile);
            throw new Exception('Cannot open temporary file "' . $tempFile . '" for writing');
        }
        stream_copy_to_stream($stream, $out);
        fclose($out);

        // touch() in makeTempFile() has already put a zero size for this path into
        // the stat cache, so it must be dropped before the size is checked
        clearstatcache(true, $tempFile);
        if (filesize($tempFile) === 0) {
            @unlink($tempFile);
            throw new Exception('The stream produced no data');
        }

        return self::openTempFile($tempFile, $options);
    }

    /**
     * Create an empty temporary file for the content of openString()/openStream()
     * and return its name
     *
     * The directory is the one used for every temporary file of the package: the
     * 'temp_dir' option, then config('fast-excel.temp_dir'), then
     * storage_path('app/tmp/fast-excel') inside Laravel, and the system temporary
     * directory outside of it.
     *
     * @param array|null $options
     *
     * @return string
     */
    protected static function makeTempFile(?array $options = []): string
    {
        $tempDir = (string)($options['temp_dir'] ?? '');
        if (!$tempDir && function_exists('config')) {
            $tempDir = (string)config('fast-excel.temp_dir');
        }
        if (!$tempDir && function_exists('storage_path')) {
            $tempDir = storage_path('app/tmp/fast-excel');
        }
        if (!$tempDir) {
            $tempDir = sys_get_temp_dir();
        }
        if (!is_dir($tempDir) && !@mkdir($tempDir, 0777, true) && !is_dir($tempDir)) {
            throw new Exception('Cannot create directory "' . $tempDir . '"');
        }
        self::cleanupTempDir($tempDir);

        $tempFile = $tempDir . DIRECTORY_SEPARATOR . uniqid('excel_reader_', true) . '.tmp';
        if (!touch($tempFile) || !is_writable($tempFile)) {
            throw new Exception('Temporary directory "' . $tempDir . '" is not writable');
        }

        return $tempFile;
    }

    /**
     * Open a temporary file created by openString()/openStream() and schedule its
     * removal
     *
     * Nothing owns this file - the reader keeps reading it while the book is alive -
     * so it is removed when the script ends, and right away if opening fails.
     *
     * @param string $tempFile
     * @param array|null $options
     *
     * @return ExcelReader
     */
    protected static function openTempFile(string $tempFile, ?array $options = []): ExcelReader
    {
        // The file has just been created and filled, so the size cached by touch()
        // must be dropped: the readers check filesize() to reject empty files
        clearstatcache(true, $tempFile);
        register_shutdown_function(static function () use ($tempFile) {
            if (is_file($tempFile)) {
                @unlink($tempFile);
            }
        });

        try {
            return self::open($tempFile, $options);
        }
        catch (\Throwable $e) {
            @unlink($tempFile);
            throw $e;
        }
    }

    /**
     * Remove stale temporary files (older than 24 hours) left after failed runs,
     * as ExcelWriter does for the same directory
     *
     * @param string $tempDir
     *
     * @return void
     */
    protected static function cleanupTempDir(string $tempDir): void
    {
        $expired = time() - 86400;
        foreach (glob($tempDir . DIRECTORY_SEPARATOR . 'excel_reader_*.tmp') ?: [] as $file) {
            if (is_file($file) && filemtime($file) < $expired) {
                @unlink($file);
            }
        }
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
