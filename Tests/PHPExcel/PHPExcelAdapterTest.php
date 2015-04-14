<?php
namespace ExcelOutput\Test;

use mbright\ExcelOutput\PHPExcel\PHPExcelAdapter;

/**
 * Test PHPExcelAdapterTest
 * @package ExcelOutput\Test
 */
class PHPExcelAdapterTest extends \PHPUnit_Framework_TestCase
{
    /** @var  PHPExcelAdapter */
    private $adapter;

    /** @var \Mockery\MockInterface */
    private $formatter;

    /** @var \Mockery\MockInterface */
    private $factory;

    public function setUp()
    {
        $this->formatter = \Mockery::mock('mbright\ExcelOutput\ExcelFormatterInterface');
        $this->factory   = \Mockery::mock('mbright\ExcelOutput\PHPExcel\PHPExcelFactory');
        $this->adapter   = new PHPExcelAdapter($this->formatter, $this->factory);
    }

    public function testExportToFormatException()
    {
        $this->setExpectedException(
            'InvalidArgumentException',
            'The given format is not supported'
        );
        $workbook = \Mockery::mock('mbright\ExcelOutput\ExcelWorkbook');
        $this->adapter->exportToFormat($workbook, 'invalidFormat');
    }

    public function testExportToFormatNoWrite()
    {
        $format   = PHPExcelAdapter::XLSX;
        $workbook = \Mockery::mock('mbright\ExcelOutput\ExcelWorkbook');
        $sheet    = \Mockery::mock('mbright\ExcelOutput\SpreadSheet');
        $PHPExcelSheet = \Mockery::mock('PHPExcel_Worksheet');
        $PHPExcelSheet->shouldIgnoreMissing();
        $phpExcel = \Mockery::mock('PHPExcel');
        $phpExcel->shouldIgnoreMissing();

        $this->factory->shouldReceive('newPHPExcel')->andReturn($phpExcel)->once();
        $workbook->shouldReceive('getMeta')->andReturn([])->twice();
        $workbook->shouldReceive('getSheets')->andReturn([$sheet]);
        $this->formatter->shouldReceive('applyRules')->with($sheet)->andReturn($PHPExcelSheet);
        $this->factory->shouldNotReceive('newWriter');

        $this->assertEquals($phpExcel, $this->adapter->exportToFormat($workbook, $format, false));
    }

    public function testExportToFormatWrite()
    {
        $format   = PHPExcelAdapter::XLSX;
        $workbook = \Mockery::mock('mbright\ExcelOutput\ExcelWorkbook');
        $sheet    = \Mockery::mock('mbright\ExcelOutput\SpreadSheet');
        $PHPExcelSheet = \Mockery::mock('PHPExcel_Worksheet');
        $PHPExcelSheet->shouldIgnoreMissing();
        $phpExcel = \Mockery::mock('PHPExcel');
        $phpExcel->shouldIgnoreMissing();
        $writer   = \Mockery::mock('\PHPExcel_Writer_Excel2007');

        $this->factory->shouldReceive('newPHPExcel')->andReturn($phpExcel)->once();
        $workbook->shouldReceive('getMeta')->andReturn([])->twice();
        $workbook->shouldReceive('getSheets')->andReturn([$sheet]);
        $this->formatter->shouldReceive('applyRules')->with($sheet)->andReturn($PHPExcelSheet);
        $this->factory->shouldReceive('newWriter')->with($format, $phpExcel)->andReturn($writer);
        $writer->shouldReceive('save')->with(null);

        $this->assertEquals($phpExcel, $this->adapter->exportToFormat($workbook, $format));
    }

    public function testFormatSheet()
    {
        $sheet = \Mockery::mock('mbright\ExcelOutput\SpreadSheet');
        $this->formatter->shouldReceive('applyRules')->with($sheet)->once();
        $this->adapter->formatSheet($sheet);
    }

    public function testHandleWorkbookMetaData()
    {
        $workbook       = \Mockery::mock('mbright\ExcelOutput\ExcelWorkbook');
        $creator        = 'creator';
        $lastModifiedBy = 'lastModBy';
        $title          = 'title';
        $subject        = 'subject';
        $description    = 'description';
        $keywords       = ['yo', 'dawg', 'foo', 'bar'];
        $categories     = ['foo', 'bar', 'yo', 'dawg'];
        $meta = [
            PHPExcelAdapter::CREATOR => $creator,
            PHPExcelAdapter::LAST_MODIFIED_BY => $lastModifiedBy,
            PHPExcelAdapter::TITLE => $title,
            PHPexcelAdapter::SUBJECT => $subject,
            PHPExcelAdapter::DESCRIPTION => $description,
            PHPExcelAdapter::KEYWORDS => $keywords,
            PHPExcelAdapter::CATEGORIES => $categories
        ];
        $workbook->shouldReceive('getMeta')->andReturn($meta);

        $phpExcel = \Mockery::Mock('PHPExcel');
        $phpExcel->shouldIgnoreMissing();
        $phpExcel->shouldReceive('getProperties->setCreator')->with($creator)->once();
        $phpExcel->shouldReceive('getProperties->setLastModifiedBy')->with($lastModifiedBy)->once();
        $phpExcel->shouldReceive('getProperties->setTitle')->with($title)->once();
        $phpExcel->shouldReceive('getProperties->setSubject')->with($subject)->once();
        $phpExcel->shouldReceive('getProperties->setDescription')->with($description)->once();
        $phpExcel->shouldReceive('getProperties->setKeywords')->with(implode(', ', $keywords))->once();
        $phpExcel->shouldReceive('getProperties->setCategory')->with(implode(', ', $categories))->once();

        $reflection = new \ReflectionClass($this->adapter);
        $method = $reflection->getMethod('handleWorkbookMetaData');
        $method->setAccessible(true);

        $this->assertEquals($phpExcel, $method->invoke($this->adapter, $workbook, $phpExcel));
    }

    public function testGetFormats()
    {
        $formats = [
            PHPExcelAdapter::XLSX,
            PHPExcelAdapter::XLS,
            PHPExcelAdapter::CSV
        ];
        $this->assertEquals($formats, $this->adapter->getFormats());
    }

    public function testSupportsFormatTrue()
    {
        $this->assertTrue($this->adapter->supportsFormat(PHPexcelAdapter::XLSX));
    }

    public function testSupportsFormatFalse()
    {
        $this->assertFalse($this->adapter->supportsFormat('TotallyWrongAndIncorrect'));
    }

    public function testSetCacheStratTrue()
    {
        $this->assertTrue(
            $this->adapter->setCacheStrategy(
                \PHPExcel_CachedObjectStorageFactory::cache_in_memory_serialized
            )
        );
    }

    public function testCacheStratFalse()
    {
        $this->assertFalse($this->adapter->setCacheStrategy('totallyWrongAndInvalid'));
    }


}
