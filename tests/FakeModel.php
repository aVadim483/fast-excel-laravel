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
}