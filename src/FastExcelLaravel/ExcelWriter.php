<?php

namespace avadim\FastExcelLaravel;

use Illuminate\Support\Collection;

class ExcelWriter  extends \avadim\FastExcelWriter\Excel
{
    /** @var array SheetWriter[] */
    protected array $sheets = [];

    /**
     * Create XLSX for export
     *
     * @param string|array $sheets
     * @param array|null $options
     *
     * @return ExcelWriter
     */
    public static function create($sheets = null, ?array $options = []): ExcelWriter
    {
        if (empty($options['temp_dir'])) {
            $tempDir = storage_path('app/tmp/fast-excel');
            if (class_exists('\File')) {
                if(!\File::isDirectory($tempDir)) {
                    \File::makeDirectory($tempDir, 0777, true, true);
                }
            }
            else {
                if (!is_dir($tempDir)) {
                    mkdir($tempDir, 0777, true);
                }
            }
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
     * @return SheetWriter
     */
    public function sheet($index = null): SheetWriter
    {
        return parent::sheet($index);
    }


    /**
     * @param $model
     * @param array|null $rowStyle
     * @param array|null $cellStyles
     *
     * @return $this
     */
    public function exportModel($model, array $rowStyle = null, array $cellStyles = null): ExcelWriter
    {
        $this->getSheet()->exportModel($model, $rowStyle, $cellStyles);

        return $this;
    }

    /**
     * @param $data
     *
     * @return $this
     */
    public function writeData($data): ExcelWriter
    {
        $this->getSheet()->writeData($data);

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
        $this->save(storage_path($filePath));

        return true;
    }

    /**
     * Store file to specified disk
     *
     * @param $disk
     * @param $path
     *
     * @return bool
     *
     * @throws \Illuminate\Contracts\Filesystem\FileExistsException
     */
    public function store($disk, $path): bool
    {
        $result = false;
        $tmpFile = $this->writer->makeTempFile();
        if ($this->writer->saveToFile($tmpFile, true, $this->getMetadata())) {
            $this->saved = true;

            $handle = fopen($tmpFile, 'rb');
            if ($handle) {
                $result = \Storage::disk($disk)->writeStream($path, $handle);
                fclose($handle);
            }
        }
        $this->writer->removeFiles();

        return $result;
    }
}
