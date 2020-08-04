<?php

use PhpOffice\PhpSpreadsheet\Calculation\ExcelException;

return [
    [
        false,
    ],
    [
        false,
        null,
    ],
    [
        false,
        -1,
    ],
    [
        false,
        0,
    ],
    [
        false,
        9,
    ],
    [
        false,
        1.5,
    ],
    [
        false,
        '',
    ],
    [
        false,
        '-1',
    ],
    [
        false,
        '2',
    ],
    [
        false,
        '-1.5',
    ],
    [
        false,
        'ABC',
    ],
    [
        true,
        ExcelException::VALUE(),
    ],
    [
        true,
        ExcelException::NA(),
    ],
    [
        false,
        'TRUE',
    ],
    [
        false,
        true,
    ],
    [
        false,
        false,
    ],
];
