<?php

use PhpOffice\PhpSpreadsheet\Calculation\ExcelException;

return [
    [
        ExcelException::VALUE(),
        ['ABC', 'DEF', 'GHI'],
    ],
    [
        ExcelException::NUM(),
        [0.0, 0.0, 0.0],
    ],
    [
        ExcelException::NUM(),
        [0.05, 10.0, -2.0],
    ],
    [
        10.33803507,
        [0.04, 10000, 15000],
    ],
    [
        3.859866162622655,
        [0.025, 2000, 2200],
    ],
    [
        87.6054764193714,
        [0.025 / 12, 1000, 1200],
    ],
    [
        3.5644379226484437,
        [0.03, 9000, 10000],
    ],
    [
        14.206699082890474,
        [0.05, 50, 100],
    ],
];
