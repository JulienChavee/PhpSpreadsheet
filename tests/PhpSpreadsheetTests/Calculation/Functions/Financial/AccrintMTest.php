<?php

namespace PhpOffice\PhpSpreadsheetTests\Calculation\Functions\Financial;

use PhpOffice\PhpSpreadsheet\Calculation\Financial;
use PhpOffice\PhpSpreadsheet\Calculation\Functions;
use PHPUnit\Framework\TestCase;

class AccrintMTest extends TestCase
{
    protected function setUp(): void
    {
        Functions::setCompatibilityMode(Functions::COMPATIBILITY_EXCEL);
    }

    /**
     * @dataProvider providerACCRINTM
     *
     * @param mixed $expectedResult
     */
    public function testACCRINTM($expectedResult, ...$args): void
    {
        $result = Financial::ACCRINTM(...$args);

        if (is_object($expectedResult)) {
            self::assertEquals($expectedResult, $result);
        } else {
            self::assertEqualsWithDelta($expectedResult, $result, 1E-8);
        }
    }

    public function providerACCRINTM()
    {
        return require 'tests/data/Calculation/Financial/ACCRINTM.php';
    }
}
