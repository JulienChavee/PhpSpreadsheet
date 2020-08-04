<?php

use PhpOffice\PhpSpreadsheet\Calculation\ExcelException;

return [
    [
        ExcelException::VALUE(),
        ['ABC', 'DEF', 'GHI'],
    ],
    [
        ExcelException::NUM(),
        [0.0, 10.0, 20.0],
    ],
    [
        ExcelException::NUM(),
        [0.05, 10.0, -2.0],
    ],
    [
        0.04137974399241062,
        [10, 10000, 15000],
    ],
    [
        0.00099330737629133,
        [96, 10000, 11000],
    ],
    [
        0.6747967875721199,
        [10, 70, 12154],
    ],
];
