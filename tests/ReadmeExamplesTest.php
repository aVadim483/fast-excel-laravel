<?php

namespace avadim\FastExcelLaravel\Test;

use avadim\FastExcelLaravel\Excel;
use avadim\FastExcelLaravel\ExcelWriter;
use avadim\FastExcelLaravel\SheetWriter;
use avadim\FastExcelReader\Excel as ExcelReader;
use Illuminate\Database\Eloquent\Model;
use Orchestra\Testbench\TestCase;
use Illuminate\Support\Collection;

class TestUser extends Model
{
    protected $fillable = ['id', 'name', 'birthday'];
    public $timestamps = false;
    
    // Fake save for testing without database
    public function save(array $options = []) { return true; }
}

class ReadmeExamplesTest extends TestCase
{
    protected string $testStorage;

    protected function setUp(): void
    {
        parent::setUp();
        $this->testStorage = __DIR__ . '/test_storage';
        if (!is_dir($this->testStorage)) {
            mkdir($this->testStorage, 0777, true);
        }
        $this->app->useStoragePath($this->testStorage);
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

    /**
     * Test export from array (README: Export Any Collections and Arrays)
     */
    public function testExportArray()
    {
        $testFileName = $this->testStorage . '/export_array.xlsx';
        if (file_exists($testFileName)) {
            unlink($testFileName);
        }

        $excel = Excel::create('Array');
        $sheet = $excel->sheet();
        
        $array = [
            [ 'id' => 1, 'name' => 'Helen' ],
            [ 'id' => 2, 'name' => 'Peter' ],
        ];
        $sheet->writeData($array);
        $excel->save($testFileName);

        $this->assertFileExists($testFileName);
        
        $reader = ExcelReader::open($testFileName);
        $rows = $reader->readRows();
        
        $this->assertEquals(1, $rows[1]['A']);
        $this->assertEquals('Helen', $rows[1]['B']);
        $this->assertEquals(2, $rows[2]['A']);
        $this->assertEquals('Peter', $rows[2]['B']);

        unlink($testFileName);
    }

    /**
     * Test export from collection (README: Export Any Collections and Arrays)
     */
    public function testExportCollection()
    {
        $testFileName = $this->testStorage . '/export_collection.xlsx';
        if (file_exists($testFileName)) {
            unlink($testFileName);
        }

        $excel = Excel::create('Collection');
        $sheet = $excel->sheet();
        
        $collection = collect([
            [ 'id' => 1, 'site' => 'google.com' ],
            [ 'id' => 2, 'site' => 'youtube.com' ],
        ]);
        $sheet->writeData($collection);
        $excel->save($testFileName);

        $this->assertFileExists($testFileName);
        
        $reader = ExcelReader::open($testFileName);
        $rows = $reader->readRows();
        
        $this->assertEquals(1, $rows[1]['A']);
        $this->assertEquals('google.com', $rows[1]['B']);

        unlink($testFileName);
    }

    /**
     * Test export with callback (README: Export Any Collections and Arrays)
     */
    public function testExportCallback()
    {
        $testFileName = $this->testStorage . '/export_callback.xlsx';
        if (file_exists($testFileName)) {
            unlink($testFileName);
        }

        $excel = Excel::create('Callback');
        $sheet = $excel->sheet();
        
        $sheet->writeData(function () {
            yield [ 'id' => 1, 'val' => 'A' ];
            yield [ 'id' => 2, 'val' => 'B' ];
        });
        $excel->save($testFileName);

        $this->assertFileExists($testFileName);
        
        $reader = ExcelReader::open($testFileName);
        $rows = $reader->readRows();
        
        $this->assertEquals('A', $rows[1]['B']);
        $this->assertEquals('B', $rows[2]['B']);

        unlink($testFileName);
    }

    /**
     * Test mapping (README: Mapping Export Data)
     */
    public function testExportMapping()
    {
        $testFileName = $this->testStorage . '/export_mapping.xlsx';
        if (file_exists($testFileName)) {
            unlink($testFileName);
        }

        $excel = Excel::create('Mapping');
        $sheet = $excel->sheet();
        
        $data = [
            (object)['id' => 1, 'first_name' => 'John', 'last_name' => 'Doe'],
            (object)['id' => 2, 'first_name' => 'Jane', 'last_name' => 'Smith'],
        ];
        
        $sheet->mapping(function($item) {
            return [
                'id' => $item->id, 
                'full_name' => $item->first_name . ' ' . $item->last_name,
            ];
        })->writeData($data);
        
        $excel->save($testFileName);

        $this->assertFileExists($testFileName);
        
        $reader = ExcelReader::open($testFileName);
        $rows = $reader->readRows();
        
        $this->assertEquals(1, $rows[1]['A']);
        $this->assertEquals('John Doe', $rows[1]['B']);

        unlink($testFileName);
    }

    /**
     * Test advanced styling (README: Advanced Usage for Data Export)
     */
    public function testExportAdvanced()
    {
        $testFileName = $this->testStorage . '/export_advanced.xlsx';
        if (file_exists($testFileName)) {
            unlink($testFileName);
        }

        $excel = Excel::create('Advanced');
        $sheet = $excel->sheet();

        $sheet->setColWidth('B', 20);
        $sheet->setColDataStyle('C', ['width' => 15, 'text-align' => 'center']);
        
        $area = $sheet->beginArea();
        $area->setValue('A1:C1', 'Header Title')
             ->applyFontStyleBold()
             ->applyTextCenter();
        
        $sheet->writeAreas();
        
        $data = [['ID', 'Name', 'Status'], [1, 'Test', 'OK']];
        $sheet->writeData($data);
        
        $excel->save($testFileName);
        $this->assertFileExists($testFileName);
        
        $reader = ExcelReader::open($testFileName);
        $rows = $reader->readRows();
        $this->assertEquals('Header Title', $rows[1]['A']);
        $this->assertEquals('ID', $rows[2]['A']);
        $this->assertEquals(1, $rows[3]['A']);

        unlink($testFileName);
    }

    /**
     * Test mapping on import (README: Mapping Import Data)
     */
    public function testImportMapping()
    {
        $testFileName = $this->testStorage . '/import_mapping.xlsx';
        
        // Create file first
        $excel = Excel::create('Import');
        $excel->sheet()->writeData([
            ['ID', 'Full Name', 'DOB'],
            [1, 'John Doe', '1990-01-01'],
        ]);
        $excel->save($testFileName);

        $excel = Excel::open($testFileName);

        $importedData = $excel->mapping(function ($row) {
            return [
                'id' => $row['A'],
                'name' => $row['B'],
                'birthday' => $row['C'],
            ];
        })->from('a2')->readRows();
        
        $this->assertCount(1, $importedData);
        $this->assertEquals('John Doe', $importedData[2]['name']);

        unlink($testFileName);
    }

    /**
     * Test reading CSV (README: Reading CSV files)
     */
    public function testReadCsv()
    {
        // importModel from a plain comma CSV (TestUser has a fake save())
        $file = $this->testStorage . '/users.csv';
        file_put_contents($file, "name,birthday\nHelen,1990-05-01\nPeter,1985-12-31\n");

        $excel = Excel::open($file);
        $result = $excel->withHeadings()->importModel(TestUser::class);
        $this->assertInstanceOf(\avadim\FastExcelLaravel\ExcelReader::class, $result);
        unlink($file);

        // CSV options passed as the second argument (semicolon + CP1251)
        $file = $this->testStorage . '/opts.csv';
        file_put_contents($file, iconv('UTF-8', 'CP1251', "name;city\nОльга;Москва\n"));

        $excel = Excel::open($file, ['delimiter' => ';', 'encoding' => 'CP1251']);
        $rows = $excel->readRows(true);
        $this->assertEquals('Ольга', $rows[2]['name']);
        $this->assertEquals('Москва', $rows[2]['city']);

        unlink($file);
    }

    /**
     * Test reading from a string and from a stream (README: Reading from a String or a Stream)
     */
    public function testReadFromStringAndStream()
    {
        $testFileName = $this->testStorage . '/from_string.xlsx';
        $excel = Excel::create('Users');
        $excel->sheet()->writeData([
            ['name', 'birthday'],
            ['Helen', '1990-05-01'],
        ]);
        $excel->save($testFileName);

        // A workbook held in a string (here: the body of a file, as Storage::get() returns it)
        $excel = Excel::openString(file_get_contents($testFileName));
        $result = $excel->withHeadings()->importModel(TestUser::class);
        $this->assertInstanceOf(\avadim\FastExcelLaravel\ExcelReader::class, $result);
        unlink($testFileName);

        // A workbook behind a stream, with CSV options
        $file = $this->testStorage . '/stream.csv';
        file_put_contents($file, "name;birthday\nHelen;1990-05-01\n");

        $stream = fopen($file, 'rb');
        $excel = Excel::openStream($stream, ['delimiter' => ';']);
        $rows = $excel->readRows(true);
        fclose($stream);

        $this->assertEquals('Helen', $rows[2]['name']);
        unlink($file);

        // The temporary copies live until the script ends, remove them right away
        $excel = null;
        gc_collect_cycles();
        foreach (glob($this->testStorage . '/app/tmp/fast-excel/excel_reader_*.tmp') ?: [] as $tempFile) {
            @unlink($tempFile);
        }
    }
}
