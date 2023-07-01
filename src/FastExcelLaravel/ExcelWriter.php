<?php

namespace avadim\FastExcelLaravel;

use Illuminate\Support\Collection;

class ExcelWriter  extends \avadim\FastExcelWriter\Excel
{
    /** @var array SheetWriter[] */
    protected array $sheets = [];

    /**
     * @param string|array $sheets
     * @param array|null $options
     *
     * @return ExcelWriter
     */
    public static function create($sheets = null, ?array $options = []): ExcelWriter
    {
        if (empty($options['temp_dir'])) {
            $tempDir = storage_path('app/tmp/fast-excel');
            if(!\File::isDirectory($tempDir)) {
                \File::makeDirectory($tempDir, 0777, true, true);
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
     * @param string $sheetName
     *
     * @return SheetWriter
     */
    public static function createSheet(string $sheetName): SheetWriter
    {
        return new SheetWriter($sheetName);
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
     * @return void
     *
     * @throws \Illuminate\Contracts\Filesystem\FileExistsException
     */
    public function store($disk, $path)
    {
        $tmpFile = $this->writer->tempFilename();
        $this->save($tmpFile);
        $handle = fopen($tmpFile, 'rb');

        \Storage::disk($disk)->writeStream($path, $handle);

        fclose($handle);
    }
}
