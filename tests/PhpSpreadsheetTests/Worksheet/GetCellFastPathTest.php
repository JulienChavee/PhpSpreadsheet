<?php

declare(strict_types=1);

namespace PhpOffice\PhpSpreadsheetTests\Worksheet;

use PhpOffice\PhpSpreadsheet\Cell\CellAddress;
use PhpOffice\PhpSpreadsheet\NamedRange;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PHPUnit\Framework\TestCase;

class GetCellFastPathTest extends TestCase
{
    public function testGetCellLowercaseLocalCoordinate(): void
    {
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->getCell('a1')->setValue('hello');

        $cell = $sheet->getCell('a1');
        self::assertSame('A1', $cell->getCoordinate());
        self::assertSame('hello', $cell->getValue());

        $spreadsheet->disconnectWorksheets();
    }

    public function testGetCellViaCellAddressObject(): void
    {
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->getCell(CellAddress::fromCellAddress('B2'))->setValue(7);

        self::assertSame(7, $sheet->getCell('B2')->getValue());

        $spreadsheet->disconnectWorksheets();
    }

    public function testGetCellViaColumnRowArray(): void
    {
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->getCell([3, 3])->setValue('via-array');

        self::assertSame('via-array', $sheet->getCell('C3')->getValue());

        $spreadsheet->disconnectWorksheets();
    }

    public function testGetCellNamedRange(): void
    {
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->getCell('D4')->setValue('named');
        $spreadsheet->addNamedRange(new NamedRange('MyCell', $sheet, '$D$4'));

        self::assertSame('named', $sheet->getCell('MyCell')->getValue());

        $spreadsheet->disconnectWorksheets();
    }
}
