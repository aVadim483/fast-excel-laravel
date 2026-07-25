<?php

namespace avadim\FastExcelLaravel\Test;


use avadim\FastExcelLaravel\Excel;
use avadim\FastExcelLaravel\ExcelWriter;
use avadim\FastExcelLaravel\SheetWriter;
use avadim\FastExcelLaravel\Test\Models\FakeModel;
use Illuminate\Support\Collection;
use avadim\FastExcelReader\Excel as ExcelReader;
use Carbon\Carbon;
use Orchestra\Testbench\TestCase;
use Illuminate\Filesystem\FilesystemManager;

class FastExcelLaravelTest extends TestCase
{
    protected ?ExcelReader $excelReader = null;
    protected array $cells = [];
    protected string $testStorage;


    protected function setUp(): void
    {
        parent::setUp();
        $this->testStorage = __DIR__ . '/test_storage';

        $this->app->useStoragePath($this->testStorage);

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
        \Schema::create('fake_models', function ($table) {
            $table->increments('id');
            $table->integer('integer')->nullable();
            $table->string('date')->nullable();
            $table->string('name')->nullable();
            $table->string('foo')->nullable();
            $table->string('bar')->nullable();
            $table->integer('int')->nullable();
            $table->timestamps();
        });
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

    protected function getStyle($cell, $flat = false): array
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
        return FakeModel::getRecords();
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

    ///////////////////////////////////////////////////////
    ///
    public function testExportArray()
    {
        $testFileName = 'test1.xlsx';
        $excel = $this->startExportTest($testFileName);

        /** @var SheetWriter $sheet */
        $sheet = $excel->getSheet();

        $data = $this->getDataArray();
        $sheet->writeData($data);

        \Config::set('filesystems.disks.dynamic', [
            'driver' => 'local',
            'root' => $this->testStorage . '/dynamic',
        ]);

        if (\Storage::disk('dynamic')->exists($testFileName)) {
            \Storage::disk('dynamic')->delete($testFileName);
        }
        $excel->store('dynamic', $testFileName);
        $path = \Storage::disk('dynamic')->path($testFileName);
        //$excel->save($testFileName);

        $this->read($path);

        $this->assertEquals(array_values($data[0]), $this->getValues('A1', 'B1', 'C1', 'D1'));

        $this->endExportTest($path);
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
        $sheet->setColStyle('c', ['width' => 12, 'text-align' => 'center']);
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


    public function testImportModelFromXls()
    {
        // Legacy XLS (Excel 97-2003) must be read through the wrapper exactly like XLSX:
        // the format is detected from the file signature, not the extension.
        $file = storage_path('test_import.xls');
        $this->assertTrue(file_exists($file));
        $this->assertTrue(ExcelReader::isXls($file));

        $excel = Excel::open($file);
        $this->assertInstanceOf(\avadim\FastExcelLaravel\ExcelReader::class, $excel);
        $this->assertInstanceOf(\avadim\FastExcelLaravel\SheetReader::class, $excel->sheet());
        $this->assertEquals('Sheet1', $excel->sheet()->name());

        // Delegated read methods work on XLS
        $rows = $excel->readRows();
        $this->assertCount(4, $rows); // header + 3 data rows
        $this->assertEquals('James Bond', $rows[2]['B']);

        // withHeadings() + importModel() work on XLS
        FakeModel::$storage = [];
        $excel->withHeadings()->importModel(FakeModel::class);
        $this->assertCount(3, FakeModel::$storage);
        $this->assertEquals('James Bond', FakeModel::$storage[0]->name);

        // mapping()->from()->readRows() chain stays "sticky" on XLS
        $excel2 = Excel::open($file);
        $mapped = $excel2->mapping(['B' => 'name'])->from('A2')->readRows();
        $this->assertEquals('James Bond', $mapped[2]['name']);
    }


    public function testExportImport()
    {
        $data = $this->getDataArray();
        $testFileName = storage_path('test_io.xlsx');

        // ** 1 ** mapping import
        FakeModel::$storage = [];
        FakeModel::query()->delete();
        foreach($data as $modelData) {
            unset($modelData['id']);
            FakeModel::create($modelData);
        }
        $excel = $this->startExportTest($testFileName);
        $sheet = $excel->getSheet();

        $sheet->exportModel(FakeModel::class);
        $excel->save($testFileName);

        $this->assertTrue(file_exists($testFileName));

        $excel = Excel::open($testFileName);
        $sheet = $excel->getSheet();
        $sheet->mapping(function ($record) {
            return [
                'integer' => $record['B'], 'date' => $record['C'], 'name' => $record['D'],
            ];
        })->importModel(FakeModel::class);
        $actual = FakeModel::storageArray();
        foreach($actual as &$row) {
            unset($row['id']);
            $row['integer'] = (int)$row['integer'];
        }
        $expected = $data;
        foreach($expected as &$row) {
            unset($row['id']);
            $row['integer'] = (int)$row['integer'];
        }
        $this->assertEquals($expected, $actual);

        // ** 2 ** mapping export/import
        FakeModel::$storage = [];
        FakeModel::query()->delete();
        foreach($data as $modelData) {
            unset($modelData['id']);
            FakeModel::create($modelData);
        }
        $excel = $this->startExportTest($testFileName);
        $sheet = $excel->getSheet();

        $sheet->mapping(function($model) {
            return [
                'integer' => $model->integer, 'date' => $model->date, 'name' => $model->name,
            ];
        })->exportModel(FakeModel::class);
        $excel->save($testFileName);

        $this->assertTrue(file_exists($testFileName));

        $excel = Excel::open($testFileName);
        $sheet = $excel->getSheet();
        $sheet->mapping(function ($record) {
            return [
                'integer' => $record['A'], 'date' => $record['B'], 'name' => $record['C'],
            ];
        })->importModel(FakeModel::class);
        $actual = FakeModel::storageArray();
        foreach($actual as &$row) {
            unset($row['id']);
            $row['integer'] = (int)$row['integer'];
        }
        $this->assertEquals($expected, $actual);

        unlink($testFileName);
    }

    public function testExportImportHead()
    {
        $data = $this->getDataArray();
        $testFileName = storage_path('test_io.xlsx');

        // ** 3 ** export/import with heading
        $excel = $this->startExportTest($testFileName);

        /** @var SheetWriter $sheet */
        $sheet = $excel->getSheet();
        $sheet->withHeadings()->mapping(function ($model) {
            return ['id' => $model->id, 'integer' => $model->integer, 'date' => Carbon::parse($model->date)->getTimestamp(), 'name' => $model->name];
        })->exportModel(FakeModel::class);
        $sheet->withHeadings()->mapping(function ($model) {
            return ['id' => $model->id, 'integer' => $model->integer, 'date' => Carbon::parse($model->date)->getTimestamp(), 'name' => $model->name];
        })->exportModel(FakeModel::class);
        $excel->save($testFileName);

        $this->assertTrue(file_exists($testFileName));

        unlink($testFileName);
    }

    public function testImportModelWithCustomHeadings()
    {
        $excel = Excel::open(storage_path('test_model.xlsx'));

        FakeModel::$storage = [];
        $excel->setDateFormat('Y-m-d');
        $excel->withHeadings(['foo', 'bar', 'int'])->importModel(FakeModel::class);
        $this->assertCount(3, FakeModel::$storage);
        $this->assertEquals('James Bond', FakeModel::$storage[0]->foo);
        $this->assertEquals(4573, FakeModel::$storage[0]->int);
        $this->assertEquals('1753-01-31', FakeModel::$storage[2]->bar);

        // custom headers are reset after import, the next one uses first row values again
        FakeModel::$storage = [];
        $excel->withHeadings()->importModel(FakeModel::class);
        $this->assertEquals('James Bond', FakeModel::$storage[0]->name);
    }

    public function testOpenWithOptions()
    {
        $tempDir = $this->testStorage . '/tmp_reader';
        $excel = Excel::open(storage_path('test_model.xlsx'), ['temp_dir' => $tempDir]);
        $this->assertEquals('Sheet1', $excel->sheet()->name());

        $prop = new \ReflectionProperty(\avadim\FastExcelReader\Reader::class, 'tempDir');
        $prop->setAccessible(true);
        $this->assertEquals(str_replace('\\', '/', $tempDir), str_replace('\\', '/', $prop->getValue()));

        \avadim\FastExcelReader\Reader::setTempDir('');
    }

    public function testSaveToReturnsResult()
    {
        $testFileName = 'test_saveto.xlsx';
        $excel = $this->startExportTest($testFileName);
        $excel->sheet()->writeData($this->getDataArray());

        $this->assertTrue($excel->saveTo($testFileName));
        $this->assertTrue(file_exists(storage_path($testFileName)));

        $this->endExportTest($testFileName);
    }

    public function testDownloadResponse()
    {
        $excel = Excel::create('Users');
        $excel->sheet()->writeData($this->getDataArray());

        $response = $excel->download('report');

        $this->assertInstanceOf(\Symfony\Component\HttpFoundation\BinaryFileResponse::class, $response);
        $disposition = $response->headers->get('Content-Disposition');
        $this->assertStringContainsString('attachment', $disposition);
        $this->assertStringContainsString('report.xlsx', $disposition);

        $file = $response->getFile()->getPathname();
        $this->assertTrue(file_exists($file));
        unlink($file);
    }

    public function testTempDirFromConfig()
    {
        $tempDir = $this->testStorage . '/tmp_config';
        config(['fast-excel.temp_dir' => $tempDir]);

        // simulate a stale temp file left from a failed run
        if (!is_dir($tempDir)) {
            mkdir($tempDir, 0777, true);
        }
        $staleFile = $tempDir . '/stale.tmp';
        file_put_contents($staleFile, 'x');
        touch($staleFile, time() - 90000);

        $testFileName = 'test_config.xlsx';
        $excel = $this->startExportTest($testFileName);
        $this->assertTrue(is_dir($tempDir));
        $this->assertFalse(file_exists($staleFile));

        $excel->sheet()->writeData($this->getDataArray());
        $this->assertTrue($excel->saveTo($testFileName));

        $this->endExportTest($testFileName);
    }

    public function testServiceProvider()
    {
        $this->app->register(\avadim\FastExcelLaravel\Providers\ExcelServiceProvider::class);

        $this->assertInstanceOf(Excel::class, $this->app->make('excel'));
        $this->assertInstanceOf(Excel::class, $this->app->make(Excel::class));
        $this->assertTrue($this->app['config']->has('fast-excel'));
        $this->assertNull(config('fast-excel.temp_dir'));
    }
}