<?php

namespace avadim\FastExcelLaravel;

use avadim\FastExcelWriter\Sheet;
use Illuminate\Support\Collection;

class SheetWriter extends Sheet
{
    private ?array $headers = null;
    private int $dataRowCount = 0;

    protected function _toArray($record)
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

    protected function _writeHeader($record)
    {
        if (!$this->headers['header_keys']) {
            $this->headers['header_keys'] = array_keys($this->_toArray($record));
        }
        if (!$this->headers['header_values']) {
            $this->headers['header_values'] = $this->headers['header_keys'];
        }

        $row = array_combine($this->headers['header_keys'], $this->headers['header_values']);
        $row = $this->headers['header_values'];
        $this->writeHeader($row, $this->headers['rowStyle'], $this->headers['colStyles']);
        ++$this->dataRowCount;
    }

    public function writeRow(array $rowValues = [], array $rowStyle = null, array $cellStyles = null): Sheet
    {
        if ($this->dataRowCount > 0 && !empty($this->headers['header_keys'])) {
            $rowData = [];
            foreach ($this->headers['header_keys'] as $key) {
                if (isset($rowValues[$key])) {
                    $rowData[] = $rowValues[$key];
                }
                else {
                    $rowData[] = null;
                }
            }
        }
        else {
            $rowData = array_values($rowValues);
        }
        return parent::writeRow($rowData, $rowStyle, $cellStyles);
    }

    /**
     * @param $data
     * @param array|null $rowStyle
     * @param array|null $cellStyles
     *
     * @return $this
     */
    public function writeData($data, array $rowStyle = null, array $cellStyles = null): SheetWriter
    {
        if (is_array($data) || ($data instanceof Collection)) {
            foreach ($data as $record) {
                if ($this->dataRowCount === 0 && $this->headers) {
                    $this->_writeHeader($record);
                }
                $this->writeRow($this->_toArray($record), $rowStyle, $cellStyles);
                ++$this->dataRowCount;
            }
        }
        elseif (is_callable($data)) {
            foreach ($data() as $record) {
                if ($this->dataRowCount === 0 && $this->headers) {
                    $this->_writeHeader($record);
                }
                $this->writeRow($this->_toArray($record), $rowStyle, $cellStyles);
                ++$this->dataRowCount;
            }
        }

        return $this;
    }

    /**
     * @param $model
     * @param array|null $rowStyle
     * @param array|null $cellStyles
     *
     * @return $this
     */
    public function exportModel($model, array $rowStyle = null, array $cellStyles = null): SheetWriter
    {
        $this->writeData(static function() use ($model) {
            foreach ($model::cursor() as $user) {
                yield $user;
            }
        }, $rowStyle, $cellStyles);

        return $this;
    }

    public function withHeaders(?array $headers = [], ?array $rowStyle = [], ?array $colStyles = []): SheetWriter
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
            'rowStyle' => $rowStyle,
            'colStyles' => $colStyles,
        ];
        $this->lastTouch['ref'] = 'row';

        return $this;
    }

}
