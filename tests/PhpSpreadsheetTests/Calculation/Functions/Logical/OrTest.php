<?php

namespace PhpOffice\PhpSpreadsheetTests\Calculation\Functions\Logical;

use PhpOffice\PhpSpreadsheet\Calculation\Functions;
use PhpOffice\PhpSpreadsheet\Calculation\Logical;
use PHPUnit\Framework\TestCase;

class OrTest extends TestCase
{
    protected function setUp(): void
    {
        Functions::setCompatibilityMode(Functions::COMPATIBILITY_EXCEL);
    }

    /**
     * @dataProvider providerOR
     *
     * @param mixed $expectedResult
     */
    public function testOR($expectedResult, ...$args): void
    {
        $result = Logical::logicalOr(...$args);

        if (is_object($expectedResult)) {
            self::assertEquals($expectedResult, $result);
        } else {
            self::assertEquals($expectedResult, $result);
        }
    }

    public function providerOR()
    {
        return require 'tests/data/Calculation/Logical/OR.php';
    }
}
