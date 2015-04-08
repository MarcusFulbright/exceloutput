<?php
namespace mbright\ExcelOutput;

/**
 * Class FormatRules
 *
 * Contains the cell range and the format rules to apply to that range.
 *
 * The exact contents of the $rules array depends on which excel library is used. PHPExcel expects a style array as
 * defined in its documentation.
 *
 * @package mbright\ExcelOutput
 */
class FormatRule
{
    /** @var null|string  */
    protected $range;

    /** @var array  */
    protected $rules;

    /**
     * @param null $range
     * @param array $rules
     */
    public function __construct($range = null, $rules = array())
    {
        $this->range = $range;
        $this->rules = $rules;
    }

    /**
     * @return null|string
     */
    public function getRange()
    {
        return $this->range;
    }

    /**
     * @param null|string $range
     */
    public function setRange($range)
    {
        $this->range = $range;
    }

    /**
     * @return array
     */
    public function getRules()
    {
        return $this->rules;
    }

    /**
     * @param array $rules
     */
    public function setRules($rules)
    {
        $this->rules = $rules;
    }
}