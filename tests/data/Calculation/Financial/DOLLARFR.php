<?php

use PhpOffice\PhpSpreadsheet\Calculation\ExcelException;

// decimal_dollar, fraction, result

return [
    [
        1.02,
        1.125,
        16,
    ],
    [
        1.04,
        1.125,
        32,
    ],
    [
        1.01,
        1.0625,
        16,
    ],
    [
        1.1000000000000001,
        1.625,
        16,
    ],
    [
        1.03,
        1.09375,
        32,
    ],
    [
        1.3,
        1.9375,
        32,
    ],
    [
        1.1200000000000001,
        1.375,
        32,
    ],
    [
        ExcelException::DIV0(),
        1.2344999999999999,
        0,
    ],
    [
        ExcelException::NUM(),
        1.2344999999999999,
        -2,
    ],
];
