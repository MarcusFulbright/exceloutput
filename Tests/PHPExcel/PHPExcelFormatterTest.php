<?php
namespace ExcelOutput\Test;

use mbright\ExcelOutput\PHPExcel\PHPExcelFormatter;
use mbright\ExcelOutput\SpreadSheet;

/**
 * Test PHPExcelFormatterTest
 * @package ExcelOutput\Test
 */
class PHPExcelFormatterTest extends \PHPUnit_Framework_TestCase
{
    /** @var \ReflectionMethod */
    private $sanitizeColumns;

    /** @var \PHPExcel_Worksheet */
    private $worksheet;

    /** @var SpreadSheet */
    private $spreadSheet;

    /** @var PHPExcelFormatter */
    private $formatter;

    private $data = [
            [1,2],
            [3,4],
            [5,6]
        ];

    public function setUp()
    {
        $this->formatter = new PHPExcelFormatter();
        $this->worksheet = new \PHPExcel_Worksheet();
        $this->spreadSheet = new SpreadSheet();
        $reflection = new \ReflectionClass($this->formatter);
        $method = $reflection->getMethod('sanitizeColumns');
        $method->setAccessible(true);
        $this->sanitizeColumns = $method;
    }

    public function testSanitizeColumnsColumnRange()
    {
        $range = 'A1:B3';
        $this->worksheet->fromArray($this->data);
        $expected = 'A1:B3';
        $result = $this->sanitizeColumns->invoke($this->formatter, $this->worksheet, $range);
        $this->assertEquals($expected, $result);
    }

    public function testSanitizeColumnRangeSingleColumn()
    {
        $range = 'A';
        $this->worksheet->fromArray($this->data);
        $expected = 'A1:A3';
        $result = $this->sanitizeColumns->invoke($this->formatter, $this->worksheet, $range);
        $this->assertEquals($expected, $result);
    }

    public function testSanitizeColumnRangeFix2ndColumn()
    {
        $range = 'A2:B';
        $this->worksheet->fromArray($this->data);
        $expected = 'A2:B3';
        $result = $this->sanitizeColumns->invoke($this->formatter, $this->worksheet, $range);
        $this->assertEquals($expected, $result);
    }

    public function testSanitizeColumnRangeFix1stColumn()
    {
        $range = 'A:B2';
        $this->worksheet->fromArray($this->data);
        $expected = 'A1:B2';
        $result = $this->sanitizeColumns->invoke($this->formatter, $this->worksheet, $range);
        $this->assertEquals($expected, $result);
    }

    public function testSanitizeColumnRangeCellValue()
    {
        $range = 'A2';
        $this->worksheet->fromArray($this->data);
        $expected = 'A2';
        $result = $this->sanitizeColumns->invoke($this->formatter, $this->worksheet, $range);
        $this->assertEquals($expected, $result);
    }
}
