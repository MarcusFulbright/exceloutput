<?php
namespace ExcelOutput\Test;

use mbright\ExcelOutput\PHPExcel\PHPExcelFactory;
use mbright\ExcelOutput\PHPExcel\PHPExcelAdapter as Adapter;

/**
 * Test PHPExcelFactoryTest
 * @package ExcelOutput\Test
 */
class PHPExcelFactoryTest extends \PHPUnit_Framework_TestCase
{
    /** @var PHPExcelFactory */
    private $factory;

    /** @var \Mockery\MockInterface */
    private $PHPExcel;

    public function setUp()
    {
        $this->factory = new PHPExcelFactory();
        $mock = \Mockery::mock('\PHPExcel');
        $mock->shouldIgnoreMissing();
        $this->PHPExcel = $mock;
    }

    public function testNewPHPExcel()
    {
        $this->assertInstanceOf('\PHPExcel', $this->factory->newPHPExcel());
    }

    public function testNewWriterExcel2007()
    {
        $this->assertInstanceOf(
            '\PHPExcel_Writer_Excel2007',
            $this->factory->newWriter(
                Adapter::XLSX,
                $this->PHPExcel
            )
        );
    }

    public function testNewWriterExcel5()
    {
        $this->assertInstanceOf('\PHPExcel_Writer_Excel5', $this->factory->newWriter(Adapter::XLS, $this->PHPExcel));
    }

    public function newWriterCSV()
    {
        $this->assertInstanceOf('PHPExcel_Writer_CSV', $this->factory->newWriter(Adapter::CSV, $this->PHPExcel));
    }
}

