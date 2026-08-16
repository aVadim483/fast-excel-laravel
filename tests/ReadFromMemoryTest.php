<?php

namespace avadim\FastExcelLaravel\Test;

use avadim\FastExcelLaravel\Excel;
use avadim\FastExcelLaravel\ExcelReader;
use avadim\FastExcelLaravel\SheetReader;
use avadim\FastExcelLaravel\Test\Models\FakeModel;
use avadim\FastExcelReader\Csv\CsvBook;
use avadim\FastExcelReader\Excel as FastExcelReader;
use avadim\FastExcelReader\Exception;
use avadim\FastExcelReader\Reader;
use avadim\FastExcelReader\Xls\XlsBook;
use Orchestra\Testbench\TestCase;

/**
 * Reading a workbook that is not a file on disk: openString() and openStream()
 */
class ReadFromMemoryTest extends TestCase
{
    protected string $testStorage;

    protected function setUp(): void
    {
        parent::setUp();
        $this->testStorage = __DIR__ . '/test_storage';
        $this->app->useStoragePath($this->testStorage);
        $this->setUpDatabase();
        FakeModel::$storage = [];
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

    protected function tearDown(): void
    {
        // The temporary files of openString()/openStream() live until the script
        // ends, so remove them here to keep the test storage clean
        Reader::setTempDir('');
        gc_collect_cycles();
        foreach ([$this->testStorage . '/app/tmp/fast-excel', $this->testStorage . '/tmp_reader'] as $dir) {
            foreach (glob($dir . '/excel_reader_*.tmp') ?: [] as $file) {
                @unlink($file);
            }
        }
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

    private function tempFiles(string $dir): array
    {
        return glob($dir . '/excel_reader_*.tmp') ?: [];
    }

    public function testOpenStringXlsx()
    {
        $content = file_get_contents($this->testStorage . '/test_model.xlsx');

        $excel = Excel::openString($content);
        $this->assertInstanceOf(ExcelReader::class, $excel);
        $this->assertInstanceOf(FastExcelReader::class, $excel->getBook());
        $this->assertInstanceOf(SheetReader::class, $excel->sheet());
        $this->assertEquals('Sheet1', $excel->sheet()->name());

        $rows = $excel->readRows();
        $this->assertCount(4, $rows); // header + 3 data rows
        $this->assertEquals('James Bond', $rows[2]['A']);
    }

    public function testOpenStringXls()
    {
        // The format is detected from the content, so a legacy XLS in a string
        // is read exactly like an XLSX one
        $content = file_get_contents($this->testStorage . '/test_import.xls');

        $excel = Excel::openString($content);
        $this->assertInstanceOf(XlsBook::class, $excel->getBook());

        $rows = $excel->readRows();
        $this->assertCount(4, $rows);
        $this->assertEquals('James Bond', $rows[2]['B']);
    }

    public function testOpenStringCsv()
    {
        // No XLSX/XLS signature: the content is read as CSV, and CSV options work
        $excel = Excel::openString("name;integer\nAnna;30\nIvan;25\n", ['delimiter' => ';']);
        $this->assertInstanceOf(CsvBook::class, $excel->getBook());

        $rows = $excel->readRows(true);
        $this->assertSame(['name' => 'Anna', 'integer' => '30'], $rows[2]);
    }

    public function testOpenStringImportModel()
    {
        $content = file_get_contents($this->testStorage . '/test_model.xlsx');

        Excel::openString($content)->withHeadings()->importModel(FakeModel::class);

        $this->assertCount(3, FakeModel::$storage);
        $this->assertEquals('James Bond', FakeModel::$storage[0]->name);
    }

    public function testOpenStringEmptyContent()
    {
        $this->expectException(Exception::class);
        Excel::openString('');
    }

    public function testOpenStreamFromFile()
    {
        $stream = fopen($this->testStorage . '/test_model.xlsx', 'rb');

        $excel = Excel::openStream($stream);
        $rows = $excel->readRows();
        $this->assertCount(4, $rows);
        $this->assertEquals('Ellen Louise Ripley', $rows[3]['A']);

        // The caller keeps the ownership of the stream, openStream() does not close it
        $this->assertTrue(is_resource($stream));
        fclose($stream);
    }

    public function testOpenStreamFromMemory()
    {
        $stream = fopen('php://memory', 'r+b');
        fwrite($stream, "name,integer\nAlice,100\nBob,200\n");
        rewind($stream);

        $rows = Excel::openStream($stream)->readRows(true);
        $this->assertSame(['name' => 'Alice', 'integer' => '100'], $rows[2]);
        fclose($stream);
    }

    public function testOpenStreamFromStorageDisk()
    {
        \Config::set('filesystems.disks.dynamic', [
            'driver' => 'local',
            'root' => $this->testStorage . '/dynamic',
        ]);
        \Storage::disk('dynamic')->put('users.csv', "name,integer\nAlice,100\n");

        $stream = \Storage::disk('dynamic')->readStream('users.csv');
        $rows = Excel::openStream($stream)->readRows(true);
        $this->assertSame(['name' => 'Alice', 'integer' => '100'], $rows[2]);

        if (is_resource($stream)) {
            fclose($stream);
        }
        \Storage::disk('dynamic')->delete('users.csv');
    }

    public function testOpenStreamNotAResource()
    {
        $this->expectException(Exception::class);
        Excel::openStream('not a stream');
    }

    public function testOpenStreamWithoutData()
    {
        $stream = fopen('php://memory', 'r+b');
        try {
            $this->expectException(Exception::class);
            Excel::openStream($stream);
        }
        finally {
            fclose($stream);
        }
    }

    public function testFacadeForwardsBothMethods()
    {
        $this->app->register(\avadim\FastExcelLaravel\Providers\ExcelServiceProvider::class);

        $excel = \avadim\FastExcelLaravel\Facades\Excel::openString("name,integer\nAlice,100\n");
        $this->assertInstanceOf(ExcelReader::class, $excel);
        $this->assertCount(2, $excel->readRows());

        $stream = fopen($this->testStorage . '/test_model.xlsx', 'rb');
        $excel = \avadim\FastExcelLaravel\Facades\Excel::openStream($stream);
        $this->assertInstanceOf(ExcelReader::class, $excel);
        $this->assertEquals('Sheet1', $excel->sheet()->name());
        fclose($stream);
    }

    public function testTempFileInDefaultDir()
    {
        $tempDir = $this->testStorage . '/app/tmp/fast-excel';
        foreach ($this->tempFiles($tempDir) as $file) {
            @unlink($file);
        }

        // No temp_dir given and no config value: storage_path('app/tmp/fast-excel')
        $excel = Excel::openString("name,integer\nAlice,100\n");
        $this->assertCount(1, $this->tempFiles($tempDir));
        $this->assertCount(2, $excel->readRows());
    }

    public function testTempFileInDirFromOptions()
    {
        $tempDir = $this->testStorage . '/tmp_reader';
        foreach ($this->tempFiles($tempDir) as $file) {
            @unlink($file);
        }

        $content = file_get_contents($this->testStorage . '/test_model.xlsx');
        $excel = Excel::openString($content, ['temp_dir' => $tempDir]);
        $this->assertCount(1, $this->tempFiles($tempDir));
        $this->assertEquals('Sheet1', $excel->sheet()->name());
    }

    public function testTempFileInDirFromConfig()
    {
        $tempDir = $this->testStorage . '/tmp_reader';
        config(['fast-excel.temp_dir' => $tempDir]);
        foreach ($this->tempFiles($tempDir) as $file) {
            @unlink($file);
        }

        // simulate a stale temp file left from a failed run
        $staleFile = $tempDir . '/excel_reader_stale.tmp';
        file_put_contents($staleFile, 'x');
        touch($staleFile, time() - 90000);

        $excel = Excel::openString("name,integer\nAlice,100\n");
        $this->assertFalse(file_exists($staleFile));
        $this->assertCount(1, $this->tempFiles($tempDir));
        $this->assertCount(2, $excel->readRows());
    }
}
