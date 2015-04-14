<?php
namespace mbright\ExcelOutput\PHPExcel;

use mbright\ExcelOutput\ExcelAdapterInterface;
use mbright\ExcelOutput\ExcelFormatterInterface;
use mbright\ExcelOutput\ExcelManagerInterface;
use mbright\ExcelOutput\ExcelWorkbook;
use mbright\ExcelOutput\SpreadSheet;

/**
 * Class PHPExcelManager
 * @package ExcelOutput\PHPExcel
 */
class PHPExcelManager implements ExcelManagerInterface
{
    /** @var ExcelAdapterInterface */
    protected $adapter;

    /** @var ExcelFormatterInterface */
    protected $formatter;

    /**
     * @param ExcelAdapterInterface $adapter
     * @param ExcelFormatterInterface $formatter
     */
    public function __construct(ExcelAdapterInterface $adapter, ExcelFormatterInterface $formatter = null)
    {
        $this->adapter   = $adapter;
        $this->formatter = $formatter;
    }

    /**
     * @param ExcelWorkbook $workbook
     * @param string $format
     * @param bool $write
     * @return mixed
     */
    public function exportToFormat(ExcelWorkbook $workbook, $format, $write = true)
    {
        return $this->adapter->exportToFormat($workbook, $format, $write);
    }

    /**
     * @param array $sheets
     * @param array $meta
     * @return ExcelWorkbook
     */
    public function newWorkbook(array $sheets, array $meta = array())
    {
        return new ExcelWorkbook($sheets, $meta);
    }

    /**
     * @param array $data
     * @param array $meta
     * @return SpreadSheet
     */
    public function newSheet(array $data, array $meta = array())
    {
        return new SpreadSheet($data, $meta);
    }

    /**
     * @param SpreadSheet $sheet
     * @throws \BadMethodCallException When the manager does not have a formatter
     * @return mixed
     */
    public function formatSheet(SpreadSheet $sheet)
    {
        if (! $this->formatter) {
            throw new \BadMethodCallException('The formatSheet Method requires the manager to have a formatter');
        }

        return $this->formatter->applyRules($sheet);
    }

    /**
     * @param ExcelWorkbook $workbook
     * @param SpreadSheet $sheet
     * @return void
     */
    public function addSheetToWorkbook(ExcelWorkbook $workbook, SpreadSheet $sheet)
    {
        $workbook->addSheet($sheet);
    }

    /**
     * @param $format
     * @return bool
     */
    public function supportsFormat($format)
    {
        return $this->adapter->supportsFormat($format);
    }
}