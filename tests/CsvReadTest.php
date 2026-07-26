<?php

namespace avadim\FastExcelLaravel\Test;

use avadim\FastExcelLaravel\ExcelReader;
use avadim\FastExcelLaravel\Test\Models\FakeModel;
use avadim\FastExcelReader\Csv\CsvBook;
use Orchestra\Testbench\TestCase;

class CsvReadTest extends TestCase
{
    protected string $testStorage;

    /** @var string[] CSV files created during a test, removed in tearDown */
    private array $created = [];

    protected function setUp(): void
    {
        parent::setUp();
        $this->testStorage = __DIR__ . '/test_storage';
        $this->app->useStoragePath($this->testStorage);
        $this->setUpDatabase();
        FakeModel::$storage = [];
    }

    protected function tearDown(): void
    {
        foreach ($this->created as $path) {
            if (is_file($path)) {
                unlink($path);
            }
        }
        $this->created = [];
        parent::tearDown();
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

    private function makeCsv(string $name, string $content): string
    {
        $path = $this->testStorage . '/' . $name;
        file_put_contents($path, $content);
        $this->created[] = $path;

        return $path;
    }

    public function testOpenDetectsCsvByContent()
    {
        // Extension is deliberately misleading; detection is by signature (none => CSV)
        $csv = $this->makeCsv('detect.txt', "name,integer\nAlice,100\n");

        $excel = ExcelReader::open($csv);
        $this->assertInstanceOf(CsvBook::class, $excel->getBook());
    }

    public function testReadRows()
    {
        $csv = $this->makeCsv('rows.csv', "name,integer,date\nAlice,100,2020-01-01\nBob,200,2020-02-02\n");

        $rows = ExcelReader::open($csv)->readRows();
        $this->assertSame(['A' => 'name', 'B' => 'integer', 'C' => 'date'], $rows[1]);
        $this->assertSame('Alice', $rows[2]['A']);

        // first row as keys
        $keyed = ExcelReader::open($csv)->readRows(true);
        $this->assertSame(['name' => 'Alice', 'integer' => '100', 'date' => '2020-01-01'], $keyed[2]);
    }

    public function testImportModel()
    {
        $csv = $this->makeCsv('import.csv', "name,integer,date\nAlice,100,2020-01-01\nBob,200,2020-02-02\n");

        ExcelReader::open($csv)->withHeadings()->importModel(FakeModel::class);

        $this->assertSame(2, FakeModel::query()->count());
        $this->assertSame(['Alice', 'Bob'], FakeModel::query()->orderBy('id')->pluck('name')->all());
        $alice = FakeModel::query()->where('name', 'Alice')->first();
        $this->assertSame(100, (int)$alice->integer);
        $this->assertSame('2020-01-01', $alice->date);
    }

    public function testImportModelWithCustomHeadings()
    {
        // No usable header row: supply attribute names in column order
        $csv = $this->makeCsv('nohead.csv', "col1,col2\nAlice,100\nBob,200\n");

        ExcelReader::open($csv)->withHeadings(['name', 'integer'])->importModel(FakeModel::class);

        // withHeadings() always skips the first row, so 2 data rows remain
        $this->assertSame(2, FakeModel::query()->count());
        $this->assertSame(['Alice', 'Bob'], FakeModel::query()->orderBy('id')->pluck('name')->all());
    }

    public function testMappingArrayAndCallback()
    {
        $csv = $this->makeCsv('map.csv', "Alice,100\nBob,200\n");

        // array remap of raw columns
        $rows = ExcelReader::open($csv)->mapping(['A' => 'name', 'B' => 'integer'])->readRows();
        $this->assertSame('Alice', $rows[1]['name']);
        $this->assertSame('100', $rows[1]['integer']);

        // callback mapping
        $rows = ExcelReader::open($csv)->mapping(function (array $row) {
            return ['name' => strtoupper($row['A']), 'integer' => (int)$row['B']];
        })->readRows();
        $this->assertSame('ALICE', $rows[1]['name']);
        $this->assertSame(100, $rows[1]['integer']);
    }

    public function testDelimiterOption()
    {
        $csv = $this->makeCsv('semi.csv', "name;integer\nAnna;30\nIvan;25\n");

        $rows = ExcelReader::open($csv, ['delimiter' => ';'])->readRows(true);
        $this->assertSame(['name' => 'Anna', 'integer' => '30'], $rows[2]);
    }

    public function testEncodingWindows1251()
    {
        $utf8 = "name,city\nОльга,Москва\n";
        $cp1251 = iconv('UTF-8', 'CP1251', $utf8);
        $csv = $this->makeCsv('cp1251.csv', $cp1251);

        $rows = ExcelReader::open($csv, ['encoding' => 'CP1251'])->readRows(true);
        $this->assertSame(['name' => 'Ольга', 'city' => 'Москва'], $rows[2]);
    }

    public function testUtf8BomIsStripped()
    {
        $csv = $this->makeCsv('bom.csv', "\xEF\xBB\xBFname,city\nOlga,Moscow\n");

        $rows = ExcelReader::open($csv)->readRows(true);
        // BOM must not leak into the first header key
        $this->assertSame('name', array_key_first($rows[2]));
        $this->assertSame('Olga', $rows[2]['name']);
    }

    public function testTolerantModeWithRaggedRows()
    {
        $csv = $this->makeCsv('ragged.csv', "# comment\nname,a,b\nX,1,2\nY,3\n\nZ,4,5,6\n");

        $rows = ExcelReader::open($csv, ['mode' => 'tolerant', 'comment_prefix' => '#'])->readRows(true);

        $this->assertSame(['name' => 'X', 'a' => '1', 'b' => '2'], $rows[3]);
        $this->assertNull($rows[4]['b']);            // short row
        $this->assertSame('6', $rows[6]['D']);        // extra column
    }

    public function testConfigCsvDefaults()
    {
        // Global default delimiter from config, no per-call option
        config()->set('fast-excel.csv', ['delimiter' => ';']);
        $csv = $this->makeCsv('cfg.csv', "name;integer\nAnna;30\n");

        $rows = ExcelReader::open($csv)->readRows(true);
        $this->assertSame(['name' => 'Anna', 'integer' => '30'], $rows[2]);

        // per-call option wins over the config default
        $csv2 = $this->makeCsv('cfg2.csv', "name,integer\nBob,40\n");
        $rows2 = ExcelReader::open($csv2, ['delimiter' => ','])->readRows(true);
        $this->assertSame(['name' => 'Bob', 'integer' => '40'], $rows2[2]);
    }
}
