<?php

namespace avadim\FastExcelLaravel;

use avadim\FastExcelWriter\Options;
use avadim\FastExcelWriter\Sheet;
use Illuminate\Support\Collection;
use Illuminate\Support\Facades\Storage;

class ExcelWriter  extends \avadim\FastExcelWriter\Excel
{
    /** @var array SheetWriter[] */
    protected array $sheets = [];

    /**
     * Create XLSX for export
     *
     * @param string|array $sheets
     * @param array|Options|null $options
     *
     * @return ExcelWriter
     */
    public static function create($sheets = null, $options = null): ExcelWriter
    {
        if (empty($options['temp_dir'])) {
            $tempDir = function_exists('config') ? config('fast-excel.temp_dir') : null;
            if (!$tempDir) {
                $tempDir = storage_path('app/tmp/fast-excel');
            }
            if (!is_dir($tempDir)) {
                mkdir($tempDir, 0777, true);
            }
            self::_cleanupTempDir($tempDir);
            if (!$options) {
                $options = ['temp_dir' => $tempDir];
            }
            else {
                $options['temp_dir'] = $tempDir;
            }
        }
        $excel = new self($options);
        if ($sheets) {
            if (is_array($sheets)) {
                foreach ($sheets as $sheetName) {
                    $excel->makeSheet($sheetName);
                }
            }
            else {
                $excel->makeSheet((string)$sheets);
            }
        }
        else {
            $excel->makeSheet();
        }

        return $excel;
    }

    /**
     * Remove stale temporary files (older than 24 hours) left after failed runs
     *
     * @param string $tempDir
     *
     * @return void
     */
    protected static function _cleanupTempDir(string $tempDir): void
    {
        $expired = time() - 86400;
        foreach (glob($tempDir . DIRECTORY_SEPARATOR . '*') ?: [] as $file) {
            if (is_file($file) && filemtime($file) < $expired) {
                @unlink($file);
            }
        }
    }

    /**
     * Create SheetWriter instance
     *
     * @param string $sheetName
     *
     * @return SheetWriter
     */
    public static function createSheet(string $sheetName): SheetWriter
    {
        return new SheetWriter($sheetName);
    }

    /**
     * Returns sheet by number or name of sheet.
     * Return the first sheet if number or name omitted
     *
     * @param int|string|null $index - number or name of sheet
     *
     * @return Sheet|null|SheetWriter
     */
    #[ \ReturnTypeWillChange]
    public function sheet($index = null): ?SheetWriter
    {
        return parent::sheet($index);
    }


    /**
     * Export a model to the current sheet
     *
     * @param $model
     * @param array|null $rowStyle
     * @param array|null $cellStyles
     *
     * @return $this
     */
    public function exportModel($model, ?array $rowStyle = null, ?array $cellStyles = null): ExcelWriter
    {
        $this->sheet()->exportModel($model, $rowStyle, $cellStyles);

        return $this;
    }

    /**
     * Write data to the current sheet
     *
     * @param $data
     *
     * @return $this
     */
    public function writeData($data): ExcelWriter
    {
        $this->sheet()->writeData($data);

        return $this;
    }

    /**
     * Save file to local storage
     *
     * @param string $filePath
     *
     * @return bool
     */
    public function saveTo(string $filePath): bool
    {
        return $this->save(storage_path($filePath));
    }

    /**
     * Store file to specified disk
     *
     * @param string $disk
     * @param string $path
     *
     * @return bool
     */
    public function store($disk, $path): bool
    {
        $result = false;
        $tmpFile = $this->writer->makeTempFile();
        if ($this->writer->saveToFile($tmpFile, true, $this->getMetadata())) {
            $this->saved = true;

            $handle = fopen($tmpFile, 'rb');
            if ($handle) {
                $result = Storage::disk($disk)->writeStream($path, $handle);
                fclose($handle);
            }
        }
        $this->writer->removeFiles();

        return $result;
    }

    /**
     * Prepare a response that downloads the generated XLSX-file
     *
     * @param string|null $name
     * @param array|null $headers
     *
     * @return \Symfony\Component\HttpFoundation\BinaryFileResponse
     */
    #[\ReturnTypeWillChange]
    public function download(?string $name = null, ?array $headers = [])
    {
        $tmpFile = $this->writer->makeTempFileName(uniqid('xlsx_'));
        $this->save($tmpFile);
        if (!$name) {
            $name = basename($tmpFile) . '.xlsx';
        }
        else {
            $name = basename($name);
            if (strtolower(pathinfo($name, PATHINFO_EXTENSION)) !== 'xlsx') {
                $name .= '.xlsx';
            }
        }

        return response()->download($tmpFile, $name, $headers ?? [])->deleteFileAfterSend(true);
    }
}
