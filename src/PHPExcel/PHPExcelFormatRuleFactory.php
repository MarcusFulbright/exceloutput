<?php
namespace ExcelOutput\PHPExcel;
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
    public function createStyleRule($range, array $style)
    {
        return new FormatRule($range, $style, 'style');
    }

    public function createNumFormatRule($range, $style)
    {
        return new FormatRule($range, $style, 'numFormat');
    }
}