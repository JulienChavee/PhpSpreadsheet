<?php

use PhpOffice\PhpSpreadsheet\Calculation\ExcelException;

return [
    [
        1800,
        [10000, 1000, 5],
    ],
    [
        1000,
        [10000, 5000, 5],
    ],
    [
        1142.857142857143,
        [8000, 0, 7],
    ],
    [
        50,
        [500, 100, 8],
    ],
    [
        2250,
        [30000, 7500, 10],
    ],
    [
        166.66666666666666,
        [1200, 200, 6],
    ],
    [
        3750,
        [45000, 7500, 10],
    ],
    [
        ExcelException::NUM(),
        [10000, 1000, -1],
    ],
    [
        ExcelException::VALUE(),
        ['INVALID', 1000, -1],
    ],
];
