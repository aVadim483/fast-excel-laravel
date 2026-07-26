<?php

return [
    /*
    |--------------------------------------------------------------------------
    | Temporary Directory
    |--------------------------------------------------------------------------
    |
    | Directory for temporary files created while reading and writing XLSX.
    | When null, storage_path('app/tmp/fast-excel') is used.
    |
    */
    'temp_dir' => null,

    /*
    |--------------------------------------------------------------------------
    | CSV Reading Options
    |--------------------------------------------------------------------------
    |
    | Default options applied when \Excel::open() reads a CSV file (a file that
    | is neither an XLSX nor a legacy XLS by its signature). These are merged
    | into the options passed to open(); a per-call option always wins, and a
    | null value here is skipped so the reader's own default is kept.
    | Options only affect CSV; XLSX and XLS ignore them.
    |
    */
    'csv' => [
        'delimiter'        => null,   // column delimiter; null = auto-detect
        'enclosure'        => '"',    // field enclosure character
        'escape'           => '',     // escape character; '' = none
        'encoding'         => null,   // source encoding, e.g. 'CP1251'; null = auto-detect (BOM is handled automatically)
        'skip_empty_lines' => true,   // skip blank lines
        'comment_prefix'   => null,   // skip lines starting with this prefix, e.g. '#'
        'mode'             => null,   // 'strict' or 'tolerant' (ragged rows); null = reader default
    ],
];
