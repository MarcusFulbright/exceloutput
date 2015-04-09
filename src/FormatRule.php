<?php
namespace mbright\ExcelOutput;

/**
 * Class FormatRules
 *
 * Contains the cell range and the format rules to apply to that range.
 *
 * The exact contents of the $rules array depends on which excel library is used. PHPExcel expects a style array as
 * defined in its documentation. The type property can optionally be used for anything that a particular adapter or
 * formatter might require.
 *
 * @package mbright\ExcelOutput
 */
class FormatRule
{
    /** @var null|string  */
    protected $range;

    /** @var array  */
    protected $rules;

    /** @var mixed */
    protected $type;

    /**
     * @param null $range
     * @param array $rules
     */
    public function __construct($range = null, array $rules = [])
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
    public function setRules(array $rules)
    {
        $this->rules = $rules;
    }

    public function getType()
    {
        return $this->type;
    }

    public function setType($type)
    {
        $this->type = $type;
    }
}