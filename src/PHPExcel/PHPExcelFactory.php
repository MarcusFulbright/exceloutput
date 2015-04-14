<?php
namespace mbright\ExcelOutput\PHPExcel;

use mbright\ExcelOutput\PHPExcel\PHPExcelAdapter as Adapter;

/**
 * Class PHPExcelFactory
 *
 * Creates PHPExcel object
 *
 * @package ExcelOutput\PHPExcel
 */
class PHPExcelFactory 
{
    /**
     * Creates a New PHPExcel Object
     *
     * @return \PHPExcel
     */
    public function newPHPExcel()
    {
        return new \PHPExcel();
    }

    /**
     * Returns the appropriate PHPExcel Writer based on the given format
     *
     * @param string $format
     * @param \PHPExcel $PHPExcel
     * @return \PHPExcel_Writer_CSV|\PHPExcel_Writer_Excel2007|\PHPExcel_Writer_Excel5
     */
    public function newWriter($format, \PHPExcel $PHPExcel)
    {
        switch (true) {
            case $format === Adapter::XLSX:
                $writer = new \PHPExcel_Writer_Excel2007($PHPExcel);
                break;
            case $format === Adapter::XLS:
                $writer = new \PHPExcel_Writer_Excel5($PHPExcel);
                break;
            case $format === Adapter::CSV:
                $writer = new \PHPExcel_Writer_CSV($PHPExcel);
                break;
            default:
                throw new \InvalidArgumentException('the given format is not supported');
        }
        return $writer;
    }
}