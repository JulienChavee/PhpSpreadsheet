<?php

use PhpOffice\PhpSpreadsheet\Calculation\ExcelException;

return [
    'Year without century specified' => [
        6890, // '11th November 1918'
        18, 11, 11,
    ],
    'Excel 1900 Calendar Base Date' => [
        1,
        1900, 1, 1,
    ],
    'Day before Excel mythical 1900 leap day' => [
        59,
        1900, 2, 28,
    ],
    'Excel mythical 1900 leap day' => [
        60,
        1900, 2, 29,
    ],
    'Day after Excel mythical 1900 leap day' => [
        61,
        1900, 3, 1,
    ],
    'Day after Excel actual 1904 leap day' => [
        713,
        1901, 12, 13,
    ],
    'signed 32-bit Unix Timestamp Earliest Date' => [
        714,
        1901, 12, 14,
    ],
    'Day before Excel 1904 Calendar Base Date' => [
        1461,
        1903, 12, 31,
    ],
    'Excel 1904 Calendar Base Date' => [
        1462,
        1904, 1, 1,
    ],
    'Day after Excel 1904 Calendar Base Date' => [
        1463,
        1904, 1, 2,
    ],
    [
        22269,
        1960, 12, 19,
    ],
    'Unix Timestamp Base Date' => [
        25569,
        1970, 1, 1,
    ],
    [
        30292,
        1982, 12, 7,
    ],
    [
        39611,
        2008, 6, 12,
    ],
    '32-bit signed Unix Timestamp Latest Date' => [
        50424,
        2038, 1, 19,
    ],
    'Day after 32-bit signed Unix Timestamp Latest Date' => [
        50425,
        2038, 1, 20,
    ],
    [
        39448,
        2008, 1, 1,
    ],
    [
        39447,
        2008, 1, null,
    ],
    [
        39446,
        2008, 1, -1,
    ],
    [
        39417,
        2008, 1, -30,
    ],
    [
        39416,
        2008, 1, -31,
    ],
    [
        39082,
        2008, 1, -365,
    ],
    [
        39508,
        2008, 3, 1,
    ],
    [
        39507,
        2008, 3, null,
    ],
    [
        39506,
        2008, 3, -1,
    ],
    [
        39142,
        2008, 3, -365,
    ],
    [
        39417,
        2008, null, 1,
    ],
    [
        39387,
        2008, -1, 1,
    ],
    [
        39083,
        2008, -11, 1,
    ],
    [
        39052,
        2008, -12, 1,
    ],
    [
        39022,
        2008, -13, 1,
    ],
    [
        39051,
        2008, -13, 30,
    ],
    [
        39021,
        2008, -13, null,
    ],
    [
        38991,
        2008, -13, -30,
    ],
    [
        38990,
        2008, -13, -31,
    ],
    [
        39814,
        2008, 13, 1,
    ],
    [
        39507,
        2007, 15, null,
    ],
    [
        40210,
        2008, 26, 1,
    ],
    [
        40199,
        2008, 26, -10,
    ],
    [
        38686,
        2008, -26, 61,
    ],
    [
        39641,
        2010, -15, -50,
    ],
    [
        39741,
        2010, -15, 50,
    ],
    [
        40552,
        2010, 15, -50,
    ],
    [
        40652,
        2010, 15, 50,
    ],
    [
        40179,
        2010, 1.5, 1,
    ],
    [
        40178,
        2010, 1.5, 0,
    ],
    [
        40148,
        2010, 0, 1.5,
    ],
    [
        40179,
        2010, 1, 1.5,
    ],
    [
        41075,
        2012, 6, 15,
    ],
    [
        41060,
        2012, 6, null,
    ],
    [
        40892,
        2012, null, 15,
    ],
    [
        167,
        null, 6, 15,
    ],
    [
        3819,
        10, 6, 15,
    ],
    [
        3622,
        10, null, null,
    ],
    [
        274,
        null, 10, null,
    ],
    [
        ExcelException::NUM(),
        null, null, 10,
    ],
    [
        ExcelException::NUM(),
        -20, null, null,
    ],
    [
        ExcelException::NUM(),
        -20, 6, 15,
    ],
    'Excel Maximum Date' => [
        2958465,
        9999, 12, 31,
    ],
    'Exceeded Excel Maximum Date' => [
        ExcelException::NUM(),
        10000, 1, 1,
    ],
    [
        39670,
        2008, 8, 10,
    ],
    [
        39813,
        2008, 12, 31,
    ],
    [
        39692,
        2008, 8, 32,
    ],
    [
        39844,
        2008, 13, 31,
    ],
    [
        39813,
        2009, 1, 0,
    ],
    [
        39812,
        2009, 1, -1,
    ],
    [
        39782,
        2009, 0, 0,
    ],
    [
        39781,
        2009, 0, -1,
    ],
    [
        39752,
        2009, -1, 0,
    ],
    [
        39751,
        2009, -1, -1,
    ],
    [
        40146,
        2010, 0, -1,
    ],
    [
        40329,
        2010, 5, 31,
    ],
    // MS Excel will fail with a #VALUE return, but PhpSpreadsheet can parse this date
    [
        40199,
        2010, 1, '21st',
    ],
    // MS Excel will fail with a #VALUE return, but PhpSpreadsheet can parse this date
    [
        40258,
        2010, 'March', '21st',
    ],
    // MS Excel will fail with a #VALUE return, but PhpSpreadsheet can parse this date
    [
        40258,
        2010, 'March', 21,
    ],
    [
        ExcelException::VALUE(),
        'ABC', 1, 21,
    ],
    [
        ExcelException::VALUE(),
        2010, 'DEF', 21,
    ],
    [
        ExcelException::VALUE(),
        2010, 3, 'GHI',
    ],
];
