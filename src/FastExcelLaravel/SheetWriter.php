<?php

namespace avadim\FastExcelLaravel;

use avadim\FastExcelWriter\Sheet;
use avadim\FastExcelWriter\Style\Style;
use Illuminate\Database\Eloquent\Builder as EloquentBuilder;
use Illuminate\Database\Eloquent\Model;
use Illuminate\Database\Eloquent\Relations\Relation;
use Illuminate\Database\Query\Builder as QueryBuilder;

class SheetWriter extends Sheet
{
    /** @var mixed|null  */
    private $mappingCallback = null;

    private array $headers = [];
    private array $attrFormats = [];
    private int $dataRowCount = 0;


    /**
     * Convert record to array
     *
     * @param $record
     *
     * @return array
     */
    protected function _toArray($record): array
    {
        if (is_object($record)) {
            if (method_exists($record, 'toArray')) {
                $result = $record->toArray();
            }
            else {
                $result = json_decode(json_encode($record), true);
            }
        }
        else {
            $result = (array)$record;
        }

        return $result;
    }

    /**
     * Write header to the sheet
     *
     * @param $record
     *
     * @return void
     */
    protected function _writeHeader($record)
    {
        if (!$this->headers['header_keys']) {
            $this->headers['header_keys'] = array_keys($this->_toArray($record));
        }
        if (!$this->headers['header_values']) {
            $this->headers['header_values'] = $this->headers['header_keys'];
        }

        //$row = array_combine($this->headers['header_keys'], $this->headers['header_values']);
        $row = $this->headers['header_values'];
        $this->writeHeader($row, $this->headers['row_style'], $this->headers['col_styles']);
        ++$this->dataRowCount;
    }

    /**
     * Write values to the current row
     *
     * @param array $rowValues
     * @param array|Style|null $rowStyle
     * @param array|null $cellStyles
     *
     * @return SheetWriter
     */
    public function writeRow(array $rowValues = [], $rowStyle = null, ?array $cellStyles = null): SheetWriter
    {
        if ($this->dataRowCount > 0 && !empty($this->headers['header_keys'])) {
            $rowData = [];
            foreach ($this->headers['header_keys'] as $key) {
                if (isset($rowValues[$key])) {
                    $rowData[$key] = $rowValues[$key];
                }
                else {
                    $rowData[] = null;
                }
            }
        }
        else {
            $rowData = $rowValues;
        }
        if ($this->attrFormats) {
            $cellStyles = (array)$cellStyles;
            foreach (array_keys($rowData) as $n => $attribute) {
                if (isset($this->attrFormats[$attribute])) {
                    $cellStyles[$n]['format'] = $this->attrFormats[$attribute];
                }
            }
        }

        return parent::writeRow($rowData, $rowStyle, $cellStyles);
    }

    /**
     * Write data to the sheet
     *
     * Accepts any iterable (array, Collection, LazyCollection, Model::cursor(), a generator, ...)
     * or a callable that returns an iterable
     *
     * @param iterable|callable $data
     * @param array|Style|null $rowStyle
     * @param array|null $colStyles
     *
     * @return $this
     */
    public function writeData($data, $rowStyle = null, ?array $colStyles = null): SheetWriter
    {
        if (is_iterable($data)) {
            $records = $data;
        }
        elseif (is_callable($data)) {
            $records = $data();
        }
        else {
            throw new \InvalidArgumentException(sprintf(
                'writeData() expects an iterable or a callable, %s given', get_debug_type($data)
            ));
        }

        foreach ($records as $record) {
            // map first: automatic headings must be taken from the keys of the mapped record,
            // because the following rows are rearranged by these keys
            if ($this->mappingCallback) {
                $record = call_user_func($this->mappingCallback, $record);
            }
            if ($this->dataRowCount === 0 && $this->headers) {
                $this->_writeHeader($record);
            }
            $this->writeRow($this->_toArray($record), $rowStyle, $colStyles);
            ++$this->dataRowCount;
        }

        return $this;
    }

    /**
     * Export a model to the sheet
     *
     * Accepts a model class name or instance (all records are exported), an Eloquent builder,
     * a query builder or a relation (only the matching records are exported). Records are read
     * lazily through cursor()
     *
     * @param string|Model|EloquentBuilder|QueryBuilder|Relation $model
     * @param array|Style|null $rowStyle
     * @param array|null $colStyles
     *
     * @return $this
     */
    public function exportModel($model, $rowStyle = null, ?array $colStyles = null): SheetWriter
    {
        if (is_string($model) || $model instanceof Model) {
            $records = $model::cursor();
        }
        elseif ($model instanceof EloquentBuilder || $model instanceof QueryBuilder || $model instanceof Relation) {
            $records = $model->cursor();
        }
        else {
            throw new \InvalidArgumentException(sprintf(
                'exportModel() expects a model class, a model, a query builder or a relation, %s given',
                get_debug_type($model)
            ));
        }
        $this->writeData($records, $rowStyle, $colStyles);
        $this->headers = [];

        return $this;
    }

    /**
     * Set headings for the sheet
     *
     * @param array|null $headers
     * @param array|null $rowStyle
     * @param array|null $colStyles
     *
     * @return $this
     */
    public function withHeadings(?array $headers = [], ?array $rowStyle = [], ?array $colStyles = []): SheetWriter
    {
        $headerKeys = $headerValues = [];
        if ($headers) {
            foreach ($headers as $key => $val) {
                if (is_string($key)) {
                    $headerKeys[] = $key;
                    $headerValues[] = $val;
                }
                else {
                    $headerKeys[] = $headerValues[] = $val;
                }
            }
        }

        $this->headers = [
            'header_keys' => $headerKeys,
            'header_values' => $headerValues,
            'row_style' => $rowStyle,
            'col_styles' => $colStyles,
        ];
        $this->lastTouch['ref'] = 'row';

        return $this;
    }

    /**
     * Set mapping callback for the sheet
     *
     * @param $callback
     *
     * @return $this
     */
    public function mapping($callback): SheetWriter
    {
        $this->mappingCallback = $callback;

        return $this;
    }

    /**
     * Set format attributes
     *
     * @param array $formats
     *
     * @return $this
     */
    public function formatAttributes(array $formats): SheetWriter
    {
        $this->attrFormats = array_replace($this->attrFormats, $formats);

        return $this;
    }
}
