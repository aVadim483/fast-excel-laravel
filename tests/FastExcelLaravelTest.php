<?php

declare(strict_types=1);

namespace avadim\FastExcelLaravel;

require_once __DIR__ . '/FakeModel.php';

use Illuminate\Support\Collection;
use avadim\FastExcelReader\Excel as ExcelReader;

final class FastExcelLaravelTest extends \Orchestra\Testbench\TestCase
{
    protected ?ExcelReader $excelReader = null;
    protected array $cells = [];
    protected string $testStorage;


    protected function setUp(): void
    {
        parent::setUp();
        $this->testStorage = __DIR__ . '/test_storage';

        app()->useStoragePath($this->testStorage);
        $this->setUpDatabase();
    }

    protected function getEnvironmentSetUp($app)
    {
        $app['config']->set('database.default', 'sqlite');
        $app['config']->set('database.connections.sqlite', [
            'driver' => 'sqlite',
            'database' => ':memory:',
            'prefix' => '',
        ]);
    }

    protected function setUpDatabase()
    {
        //$res = $this->artisan('migrate')->run();
        //$this->loadMigrationsFrom(__DIR__ . '/../database/migrations/000_create_test_models_table.php');
    }


    protected function getValue($cell)
    {
        preg_match('/^(\w+)(\d+)$/', strtoupper($cell), $m);

        return $this->cells[$m[2]][$m[1]]['v'] ?? null;
    }

    protected function getValues(...$cells): array
    {
        $result = [];
        foreach ($cells as $cell) {
            $result[] = $this->getValue($cell);
        }

        return $result;
    }

    protected function getStyle($cell, $flat = false)
    {
        preg_match('/^(\w+)(\d+)$/', strtoupper($cell), $m);
        $styleIdx = $this->cells[$m[2]][$m[1]]['s'] ?? null;
        if ($styleIdx !== null) {
            $style = $this->excelReader->getCompleteStyleByIdx($styleIdx);
            if ($flat) {
                $result = [];
                foreach ($style as $key => $val) {
                    $result = array_merge($result, $val);
                }
            }
            else {
                $result = $style;
            }

            return $result;
        }

        return [];
    }


    protected function getDataArray(): array
    {
        return include __DIR__ . '/FakeData.php';
    }

    protected function getDataCollectionStd(): Collection
    {
        $data = $this->getDataArray();
        $result = [];
        foreach ($data as $row) {
            $result[] = (object)$row;
        }

        return collect($result);
    }

    protected function read($testFileName)
    {
        $this->assertTrue(file_exists($testFileName));

        $this->excelReader = ExcelReader::open($testFileName);
        $this->cells = $this->excelReader->readRows(false, null, true);
    }


    protected function startExportTest($testFileName, $sheets = []): ExcelWriter
    {
        if (file_exists($testFileName)) {
            unlink($testFileName);
        }
        elseif (file_exists(storage_path($testFileName))) {
            unlink(storage_path($testFileName));
        }
        FakeModel::$storage = [];

        return Excel::create($sheets);
    }

    protected function endExportTest($testFileName)
    {
        $this->excelReader = null;
        $this->cells = [];

        if (file_exists($testFileName)) {
            unlink($testFileName);
        }
        elseif (file_exists(storage_path($testFileName))) {
            unlink(storage_path($testFileName));
        }
    }

    public function testExportArray()
    {
        $testFileName = __DIR__ . '/test1.xlsx';
        $excel = $this->startExportTest($testFileName);

        /** @var SheetWriter $sheet */
        $sheet = $excel->getSheet();

        $data = $this->getDataArray();
        $sheet->writeData($data);
        $excel->save($testFileName);

        $this->read($testFileName);

        $this->assertEquals(array_values($data[0]), $this->getValues('A1', 'B1', 'C1', 'D1'));

        $this->endExportTest($testFileName);
    }

    public function testExportArrayWithHeaders()
    {
        $testFileName = __DIR__ . '/test2.xlsx';
        $excel = $this->startExportTest($testFileName);

        /** @var SheetWriter $sheet */
        $sheet = $excel->getSheet();

        $data = $this->getDataArray();
        $sheet->withHeadings()->writeData($data);
        $excel->save($testFileName);

        $this->read($testFileName);
        $row = $data[1];

        $this->assertEquals(array_keys($row), $this->getValues('A1', 'B1', 'C1', 'D1'));
        $this->assertEquals(array_values($row), $this->getValues('A3', 'B3', 'C3', 'D3'));

        $this->endExportTest($testFileName);
    }

    public function testExportCollection()
    {
        $testFileName = __DIR__ . '/test3.xlsx';
        $excel = $this->startExportTest($testFileName);

        /** @var SheetWriter $sheet */
        $sheet = $excel->getSheet();

        $data = $this->getDataArray();
        $sheet->writeData(collect($this->getDataCollectionStd()));
        $excel->save($testFileName);

        $this->read($testFileName);

        $this->assertEquals(array_values($data[0]), $this->getValues('A1', 'B1', 'C1', 'D1'));

        $this->endExportTest($testFileName);
    }

    public function testExportCollectionWithHeaders()
    {
        $testFileName = 'test4.xlsx';
        $excel = $this->startExportTest($testFileName);

        /** @var SheetWriter $sheet */
        $sheet = $excel->getSheet();

        $sheet->withHeadings(['date', 'name'])
            ->applyFontStyleBold()
            ->applyBorder('thin')
            ->writeData(collect($this->getDataCollectionStd()));
        $excel->saveTo($testFileName);

        $this->read(storage_path($testFileName));

        $this->assertEquals(['1753-01-31', 'Captain Jack Sparrow', null, null], $this->getValues('A4', 'B4', 'C4', 'D4'));

        $this->endExportTest($testFileName);
    }

    public function testExportMultipleSheets()
    {
        $testFileName = 'test5.xlsx';
        $excel = $this->startExportTest($testFileName);

        $sheet = $excel->makeSheet('Collection');
        $collection = collect([
            [ 'id' => 1, 'site' => 'google.com' ],
            [ 'id' => 2, 'site.com' => 'youtube.com' ],
        ]);
        $sheet->writeData($collection);

        $sheet = $excel->makeSheet('Array');
        $array = [
            [ 'id' => 1, 'name' => 'Helen' ],
            [ 'id' => 2, 'name' => 'Peter' ],
        ];
        $sheet->writeData($array);

        $sheet = $excel->makeSheet('Callback');
        $sheet->writeData(function () {
            for ($i = 1; $i <= 3; $i++) {
                yield [$i, $i * 2, $i * 3];
            }
        });

        $excel->saveTo($testFileName);
        $file = storage_path($testFileName);

        $this->assertTrue(file_exists($file));

        $this->excelReader = ExcelReader::open($file);
        $this->excelReader->selectSheet('Collection');
        $this->cells = $this->excelReader->readRows(false, null, true);
        $this->assertEquals('youtube.com', $this->getValue('b2'));

        $this->excelReader->selectSheet('Array');
        $this->cells = $this->excelReader->readRows(false, null, true);
        $this->assertEquals('Peter', $this->getValue('b2'));

        $this->excelReader->selectSheet('Callback');
        $this->cells = $this->excelReader->readRows(false, null, true);
        $this->assertEquals(9, $this->getValue('C3'));

        $this->endExportTest($testFileName);
    }

    public function testExportAdvanced()
    {
        $testFileName = 'test6.xlsx';
        $excel = $this->startExportTest($testFileName);

        /** @var SheetWriter $sheet */
        $sheet = $excel->getSheet();

        $sheet->setColWidth('B', 12);
        $sheet->setColOptions('c', ['width' => 12, 'text-align' => 'center']);
        $sheet->setColWidth('d', 'auto');

        $title = 'This is demo of avadim/fast-excel-laravel';
        $area = $sheet->beginArea();
        $area->setValue('A2:D2', $title)
            ->applyFontSize(14)
            ->applyFontStyleBold()
            ->applyTextCenter();

        $area
            ->setValue('a4:a5', '#')
            ->setValue('b4:b5', 'Number')
            ->setValue('c4:d4', 'Movie Character')
            ->setValue('c5', 'Birthday')
            ->setValue('d5', 'Name')
        ;
        $area->withRange('a4:d5')
            ->applyBgColor('#ccc')
            ->applyFontStyleBold()
            ->applyOuterBorder('thin')
            ->applyInnerBorder('thick')
            ->applyTextCenter();
        $sheet->writeAreas();

        $sheet->writeData(collect($this->getDataCollectionStd()));
        $excel->saveTo($testFileName);

        $this->read(storage_path($testFileName));

        $this->assertEquals([982630, '2179-08-12', 'Ellen Louise Ripley', null], $this->getValues('B7', 'C7', 'D7', 'e7'));

        $this->endExportTest($testFileName);
    }

    public function testImportModel()
    {
        $testFileName = 'test_model.xlsx';
        $excel = Excel::open(storage_path($testFileName));
        $this->assertEquals('Sheet1', $excel->sheet()->name());

        FakeModel::$storage = [];
        $excel->withHeadings()->importModel(FakeModel::class);
        $this->assertCount(3, FakeModel::$storage);
        $this->assertEquals('James Bond', FakeModel::$storage[0]->name);

        FakeModel::$storage = [];
        $excel->setDateFormat('Y-m-d');
        $excel->mapping(['A' => 'foo', 'B' => 'bar', 'C' => 'int'])->importModel(FakeModel::class, 'B4');
        $this->assertEquals('1753-01-31', FakeModel::$storage[0]->bar);

        $testFileName = 'test_model2.xlsx';
        $excel = Excel::open(storage_path($testFileName));

        FakeModel::$storage = [];
        $excel->withHeadings()->importModel(FakeModel::class, 'b4');
        $this->assertCount(3, FakeModel::$storage);
        $this->assertEquals('James Bond', FakeModel::$storage[0]->name);

        FakeModel::$storage = [];
        $excel->importModel(FakeModel::class, 'b5:d5', ['B' => 'foo', 'C' => 'bar', 'D' => 'int']);
        $this->assertCount(1, FakeModel::$storage);
        $this->assertEquals('James Bond', FakeModel::$storage[0]->foo);
        $this->assertFalse(isset(FakeModel::$storage[1]));

        FakeModel::$storage = [];
        $excel->setDateFormat('Y-m-d');
        $excel->importModel(FakeModel::class, 'b5', ['B' => 'foo', 'C' => 'bar', 'D' => 'int']);
        $this->assertCount(3, FakeModel::$storage);
        $this->assertEquals('Captain Jack Sparrow', FakeModel::$storage[2]->foo);
        $this->assertEquals('1753-01-31', FakeModel::$storage[2]->bar);
        $this->assertEquals(7239, FakeModel::$storage[2]->int);

        $sheet = $excel->sheet();
        $sheet->setReadArea('b5');
        $result = [];
        foreach ($sheet->nextRow() as $rowNum => $rowData) {
            $result[$rowNum] = $rowData;
        }
        $this->assertCount(3, $result);
        $this->assertEquals('James Bond', $result[5]['B']);
        $this->assertEquals('Ellen Louise Ripley', $result[6]['B']);
        $this->assertEquals('Captain Jack Sparrow', $result[7]['B']);
    }


    public function testExportImport()
    {
        $data = $this->getDataArray();
        $testFileName = __DIR__ . '/test_io.xlsx';

        // ** 1 ** mapping import
        FakeModel::$storage = [];
        $excel = $this->startExportTest($testFileName);
        $sheet = $excel->getSheet();

        $sheet->exportModel(FakeModel::class);
        $excel->save($testFileName);

        $this->assertTrue(file_exists($testFileName));

        $excel = Excel::open($testFileName);
        $sheet = $excel->getSheet();
        $sheet->mapping(function ($record) {
            return [
                'id' => $record['A'], 'integer' => $record['B'], 'date' => $record['C'], 'name' => $record['D'],
            ];
        })->importModel(FakeModel::class);
        $this->assertEquals($data, FakeModel::storageArray());

        // ** 2 ** mapping export/import
        FakeModel::$storage = [];
        $excel = $this->startExportTest($testFileName);
        $sheet = $excel->getSheet();

        $sheet->mapping(function($model) {
            return [
                'id' => $model->id, 'integer' => $model->integer, 'date' => $model->date, 'name' => $model->name,
            ];
        })->exportModel(FakeModel::class);
        $excel->save($testFileName);

        $this->assertTrue(file_exists($testFileName));

        $excel = Excel::open($testFileName);
        $sheet = $excel->getSheet();
        $sheet->mapping(function ($record) {
            return [
                'id' => $record['A'], 'integer' => $record['B'], 'date' => $record['C'], 'name' => $record['D'],
            ];
        })->importModel(FakeModel::class);
        $this->assertEquals($data, FakeModel::storageArray());
    }

    public function testExportImportHead()
    {
        $data = $this->getDataArray();
        $testFileName = __DIR__ . '/test_io.xlsx';

        // ** 3 ** export/import with heading
        $excel = $this->startExportTest($testFileName);

        $sheet = $excel->getSheet();
        $sheet->withHeadings()->exportModel(FakeModel::class);
        $excel->save($testFileName);

        $this->assertTrue(file_exists($testFileName));

        $excel = Excel::open($testFileName);
        $sheet = $excel->getSheet();
        $sheet->withHeadings()->importModel(FakeModel::class);
        $this->assertEquals($data, FakeModel::storageArray());

        // ** 4 ** format dates
        $excel = $this->startExportTest($testFileName);

        $sheet = $excel->getSheet();
        $sheet->withHeadings()->setFieldFormats(['date' => '@date'])->exportModel(FakeModel::class);
        $excel->save($testFileName);

        $this->assertTrue(file_exists($testFileName));

        $excel = Excel::open($testFileName);
        $excel->setDateFormat('Y-m-d');
        $sheet = $excel->getSheet();
        $sheet->withHeadings()->importModel(FakeModel::class);
        $this->assertEquals($data, FakeModel::storageArray());

        $this->endExportTest($testFileName);
    }

}