<?php
namespace mbright\ExcelOutput;

/**
 * Class FormatRule
 *
 * Used to store a set of format rules
 *
 * @package mbright\ExcelOutput
 */
class FormatRule
{
    /**
     * @var string Range to apply the rule to
     */
    protected $range;

    /**
     * @var string Rule to apply
     */
    protected $rule;

    /**
     * @param string $range
     * @param string $rule
     */
    public function __construct($range = null, $rule = null)
    {
        if ($range != null && $rule != null) {
            $this->range = $range;
            $this->rule = $rule;
        }
    }

    /**
     * @return null|string
     */
    public function getRange()
    {
        return $this->range;
    }

    /**
     * @param $range
     */
    public function setRange($range)
    {
        $this->range = $range;
    }

    /**
     * @return null|string
     */
    public function getRule()
    {
        return $this->rule;
    }

    /**
     * @param $rule
     */
    public function setRule($rule)
    {
        $this->rule = $rule;
    }
}