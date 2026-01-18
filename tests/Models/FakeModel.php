<?php

namespace avadim\FastExcelLaravel\Test\Models;

use Illuminate\Database\Eloquent\Model;

class FakeModel extends Model
{
    public static $storage = [];

    protected $fillable = ['id', 'integer', 'date', 'name', 'foo', 'bar', 'int'];

    public function __construct(array $attributes = [])
    {
        parent::__construct($attributes);
    }

    public function save(array $options = [])
    {
        self::$storage[] = $this;
        return parent::save($options);
    }

    public static function getRecords()
    {
        return [
            ['id' => 1, 'integer' => 982630, 'date' => '2179-08-12', 'name' => 'James Bond'],
            ['id' => 2, 'integer' => 982630, 'date' => '2179-08-12', 'name' => 'Ellen Louise Ripley'],
            ['id' => 3, 'integer' => 300, 'date' => '1753-01-31', 'name' => 'Captain Jack Sparrow'],
        ];
    }

    public static function storageArray()
    {
        $result = [];
        foreach (self::$storage as $model) {
            $data = $model->toArray();
            unset($data['updated_at'], $data['created_at']);
            $result[] = $data;
        }
        return $result;
    }

    public static function all($columns = ['*'])
    {
        $records = self::getRecords();
        $collection = collect();
        foreach ($records as $record) {
            $collection->push(new self($record));
        }
        return $collection;
    }
}
