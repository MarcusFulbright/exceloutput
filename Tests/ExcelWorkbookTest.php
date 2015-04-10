<?php
namespace ExcelOutput\Test;
use mbright\ExcelOutput\ExcelWorkbook;

/**
 * Test ExcelWorkbookTest
 * @package ExcelOutput\Test
 */
class ExcelWorkbookTest extends \PHPUnit_Framework_TestCase
{
    /** @var array */
    private $sheets = [];

    /** @var array */
    private $meta = ['meta'];

    /** @var ExcelWorkbook */
    private $workbook;

    public function setUp()
    {
        $sheetMock = \Mockery::mock('mbright\ExcelOutput\SpreadSheet');
        $this->sheets[] = $sheetMock;

        $this->workbook = new ExcelWorkbook($this->sheets, $this->meta);
    }

    public function testConstructPopulated()
    {
        $this->assertEquals($this->sheets, $this->workbook->getSheets());
        $this->assertEquals($this->meta, $this->workbook->getMeta());
    }

    public function testConstructEmpty()
    {
        $workbook = new ExcelWorkbook();
        $this->assertInstanceOf('mbright\ExcelOutput\ExcelWorkbook', $workbook);
        $this->assertEmpty($workbook->getSheets());
        $this->assertEmpty($workbook->getMeta());
    }

    public function testValidateSheetsTrue()
    {
        $relfection = new \ReflectionClass($this->workbook);
        $method = $relfection->getMethod('validateSheets');
        $method->setAccessible(true);
        $this->assertTrue($method->invoke($this->workbook, $this->sheets));
    }

    public function testValidateSheetsFalse()
    {
        $relfection = new \ReflectionClass($this->workbook);
        $method = $relfection->getMethod('validateSheets');
        $method->setAccessible(true);
        $this->assertFalse($method->invoke($this->workbook, ['invalid']));
    }

    public function testConstructThrowsException()
    {
        $this->setExpectedException(
            'InvalidArgumentException',
            '$sheets can only contain SpreadSheet objects'
        );
        new ExcelWorkbook(['invalid']);
    }

    public function testGetSheets()
    {
        $this->assertEquals($this->sheets, $this->workbook->getSheets());
    }

    public function testAddSheet()
    {
        $newSheet = \Mockery::mock('mbright\ExcelOutput\SpreadSheet');
        $this->workbook->addSheet($newSheet);
        $this->sheets[] = $newSheet;
        $this->assertEquals($this->sheets, $this->workbook->getSheets());
    }

    public function testGetMeta()
    {
        $this->assertEquals($this->meta, $this->workbook->getMeta());
    }

    public function testSetMeta()
    {
        $meta = ['newMeta'];
        $this->workbook->setMeta($meta);
        $this->assertEquals($meta, $this->workbook->getMeta());
    }
}
