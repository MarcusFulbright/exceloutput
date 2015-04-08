<?php
namespace ExcelOutput\PHPExcel;
use mbright\ExcelOutput\ExcelFormatterInterface;
use mbright\ExcelOutput\SpreadSheet;
use mbright\ExcelOutput\FormatRule;

/**
 * Class PHPExcelFormatter
 * @package ExcelOutput\PHPExcel
 */
class PHPExcelFormatter implements  ExcelFormatterInterface
{
    /**
     * @param SpreadSheet $sheet
     * @param null|\PHPExcel_Worksheet $target
     * @return \PHPExcel_Worksheet
     */
    public function applyRules(SpreadSheet $sheet, $target = null)
    {
        $title = (isset($sheet->getMeta()[PHPExcelAdapter::TITLE])) ? $sheet->getMeta()[PHPExcelAdapter::TITLE] : null;
        $target = new \PHPExcel_Worksheet($title);
        $target->fromArray($sheet->getData());

        /** @var FormatRule $rule */
        foreach ($sheet->getFormatRules() as $rule) {
            $range = $this->sanitizeColumns($target, $rule->getRange());
            $target->getStyle($range)->applyFromArray($rule->getRules());
        }
        return $target;
    }

    /**
     * Converts A:B Column declaration to include the correct cell values for PHPExcel.
     *
     * @param \PHPExcel_Worksheet $sheet
     * @param $range
     * @throws \Exception
     * @return string
     */
    private function sanitizeColumns(\PHPExcel_Worksheet $sheet, $range)
    {
        //Should be able to handle B and B:N and still work with B7:N5.
        if (preg_match("/^[a-z]+[\d]*($|:([a-z]+[\d]*$))/i", $range) === 0) {
            throw new \Exception($range . " is not a parseable columns selection.");
        }
        $columnarray = explode(':', $range);
        if (preg_match("/[\d]/", $columnarray[0]) === 0) {
            $columnarray[0] .= "1";
        }
        if (isset($columnarray[1]) === false) {
            $columnarray[1] = preg_replace("/[\d]/", "", $columnarray[0]);
        }
        if (preg_match("/[\d]/", $columnarray[1]) === 0) {
            $columnarray[1] .= $sheet->getHighestRow();
        }
        return implode(':', $columnarray);
    }
}