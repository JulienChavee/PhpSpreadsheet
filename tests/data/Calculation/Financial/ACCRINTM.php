<?php

use PhpOffice\PhpSpreadsheet\Calculation\ExcelException;

// Issue date, Settlement, Rate, Par, Basis, Result

return [
    [
        20.547945205478999,
        '2008-04-01',
        '2008-06-15',
        0.10000000000000001,
        1000,
        3,
    ],
    [
        800,
        '2010-01-01',
        '2010-12-31',
        0.080000000000000002,
        10000,
    ],
    [
        ExcelException::NUM(),
        '2008-03-05',
        '2008-08-31',
        -0.10000000000000001,
        1000,
        2,
    ],
    [
        ExcelException::VALUE(),
        'Invalid Date',
        '2008-08-31',
        0.10000000000000001,
        1000,
        2,
    ],
    [
        ExcelException::VALUE(),
        '2008-03-01',
        '2008-08-31',
        'ABC',
        1000,
        2,
    ],
];
