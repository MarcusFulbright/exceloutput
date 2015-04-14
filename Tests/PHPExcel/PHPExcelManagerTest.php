<?php
namespace ExcelOutput\Test;

use mbright\ExcelOutput\PHPExcel\PHPExcelManager;

/**
 * Test PHPExcelManagerTest
 * @package ExcelOutput\Test
 */
class PHPExcelManagerTest extends \PHPUnit_Framework_TestCase
{
    /** @var \Mockery\MockInterface */
    private $adapter;

    /** @var \Mockery\MockInterface */
    private $formatter;

    /** @var PHPExcelManager */
    private $manager;

    public function setUp()
    {
        $this->adapter   = \Mockery::mock('mbright\ExcelOutput\ExcelAdapterInterface');
        $this->formatter = \Mockery::mock('mbright\ExcelOutput\ExcelFormatterInterface');
        $this->manager = new PHPExcelManager($this->adapter, $this->formatter);
    }

    public function testExportToFormat()
    {
        $workbook = \Mockery::mock('mbright\ExcelOutput\ExcelWorkbook');
        $format = 'format';
        $this->adapter->shouldReceive('exportToFormat')->with($workbook, $format, true);
        $this->manager->exportToFormat($workbook, $format);
    }

    public function testNewWorkbook()
    {
        $this->assertInstanceof('mbright\ExcelOutput\ExcelWorkbook', $this->manager->newWOrkbook([]));
    }

    public function testNewSheet()
    {
         $this->assertInstanceof('mbright\ExcelOutput\SpreadSheet', $this->manager->newSheet([]));
    }

    public function testFormatSheet()
    {
        $sheet = \Mockery::mock('mbright\ExcelOutput\SpreadSheet');
        $this->formatter->shouldReceive('applyRules')->with($sheet);
        $this->manager->formatSheet($sheet);
    }

    public function testAddSheetToWorkBook()
    {
        $workbook = \Mockery::mock('mbright\ExcelOutput\ExcelWorkbook');
        $sheet = \Mockery::mock('mbright\ExcelOutput\SpreadSheet');
        $workbook->shouldReceive('addSheet')->with($sheet);
        $this->manager->addSheetToWorkbook($workbook, $sheet);
    }

    public function testSupportsFormat()
    {
        $format = 'format';
        $this->adapter->shouldReceive('supportsFormat')->with($format);
        $this->manager->supportsFormat($format);
    }
}
