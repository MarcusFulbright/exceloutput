<?php
namespace mbright\ExcelOutput\PHPExcel;

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
            switch (true) {
                case ($rule->getType() === 'style'):
                    $target->getStyle($range)->applyFromArray($rule->getRules());
                    break;
                case ($rule->getType() === 'numFormat'):
                    $target->getStyle($range)->getNumberFormat()->setFormatCode($rule->getRules()[0]);
            }
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
        $columns = explode(':', $range);
        $count = count($columns);
        $regex = "/^\D+\d+$/";
        if ($count == 1 && preg_match($regex, $columns[0]) === 0) {
            $columns[1] = $columns[0];
        }
        foreach ($columns as $position => &$column) {
            $hasInt = preg_match($regex, $column) === 1;
            $needsFirstCell = $position === 0 && ! $hasInt;
            $needsLastCell = $position === 1 && ! $hasInt;
            switch (true) {
                case $needsFirstCell:
                    $column .= '1';
                    break;
                case $needsLastCell:
                    $column .= $sheet->getHighestRow($column);
                    break;
            }
        }
        return implode(':', $columns);
    }
}