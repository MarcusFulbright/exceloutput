<?php
namespace mbright\ExcelOutput\PHPExcel;

use mbright\ExcelOutput\FormatRule;

/**
 * Class FormatRuleFactory
 *
 * Used to create format rules that comply with PHPExcel
 *
 * @package ExcelOutput\PHPExcel
 */
class PHPExcelFormatRuleFactory
{
    /**
     * Creates a format rule that PHPExcel can use to apply a $styleArray.
     *
     * @see PHPExcel documentation for the composition of a $styleArray
     * @param $range
     * @param array $rules
     * @return FormatRule
     */
    public function createStyleRule($range, array $rules)
    {
        $rule =  new FormatRule($range, $rules);
        $rule->setType('style');
        return $rule;
    }

    /**
     * Creates a format rule that PHPExcel can use to apply a number format
     *
     * @see PHPExcel documentation for valid numberformat constants
     * @param $range
     * @param $rule
     * @return FormatRule
     */
    public function createNumFormatRule($range, array $rule)
    {
        if (count($rule) > 1) {
            throw new \InvalidArgumentException('PHPExcel can only take 1 item as a number format');
        }

        $rule =  new FormatRule($range, $rule, 'numFormat');
        $rule->setType('numFormat');
        return $rule;
    }
}