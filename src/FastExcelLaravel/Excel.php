<?php

namespace avadim\FastExcelLaravel;

class Excel
{
    /**
     * Excel constructor
     *
     * @param array|null $options
     */
    public function __construct(?array $options = [])
    {
    }

    /**
     * Create new XLSX-file for export
     *
     * @param array|string|null $sheets
     * @param array|null $options
     *
     * @return ExcelWriter
     */
    public static function create($sheets = null, ?array $options = []): ExcelWriter
    {
        return ExcelWriter::create($sheets, $options);
    }

    /**
     * Open an existing XLSX-file for import
     *
     * @param string $file
     * @param array|null $options
     *
     * @return ExcelReader
     */
    public static function open(string $file, ?array $options = []): ExcelReader
    {
        return ExcelReader::open($file, $options);
    }

    /**
     * Open a workbook held in a string for import
     *
     * @param string $content
     * @param array|null $options
     *
     * @return ExcelReader
     */
    public static function openString(string $content, ?array $options = []): ExcelReader
    {
        return ExcelReader::openString($content, $options);
    }

    /**
     * Open a workbook from an open stream resource for import
     *
     * @param resource $stream
     * @param array|null $options
     *
     * @return ExcelReader
     */
    public static function openStream($stream, ?array $options = []): ExcelReader
    {
        return ExcelReader::openStream($stream, $options);
    }
}
