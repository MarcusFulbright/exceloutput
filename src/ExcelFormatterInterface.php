<?php
namespace mbright\ExcelOutput;

/**
 * Interface ExcelFormatterInterface
 *
 * Optional class that can contain logic required to apply formatting rules.
 *
 * In some cases, the logic reburied to apply format rules is simply too big or complex to live in a manager or adapter.
 * For those instances, a special formatter class can be used. The composition of these classes will really depend on
 * the library you're building the adapter for.
 *
 * @package mbright\ExcelOutput
 */
interface ExcelFormatterInterface
{
    /**
     * Applies the format rules associated with the given SpreadSheet.
     *
     * In some cases, you will need to apply the format rules to a particular target. In those cases, a $target can get
     * passed in.
     *
     * @param SpreadSheet $sheet
     * @param null $target
     * @return mixed
     */
    public function applyRules(SpreadSheet $sheet, $target = null);
}