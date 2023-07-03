<?php

namespace avadim\FastExcelLaravel;

class FakeModel
{
    public static array $storage = [];

    protected array $attributes = [];
    protected array $fillable = [];

    public function __construct(array $attributes = [])
    {
        $this->fill($attributes);
    }

    public static function create(array $attributes = []): FakeModel
    {
        return new self($attributes);
    }

    public static function cursor(): \Generator
    {
        $data = include __DIR__ . '/FakeData.php';
        foreach ($data as $record) {
            yield new self($record);
        }
    }

    public function fill(array $attributes): FakeModel
    {
        $this->attributes = $attributes;

        return $this;
    }

    public function save(array $options = []): bool
    {
        self::$storage[] = $this;

        return true;
    }

    public function __get($name)
    {
        return $this->attributes[$name] ?? null;
    }

    public function toArray(): array
    {
        return $this->attributes;
    }

    public static function storageArray(): array
    {
        $result = [];
        foreach (FakeModel::$storage as $item) {
            $result[] = $item->toArray();
        }

        return $result;
    }
}