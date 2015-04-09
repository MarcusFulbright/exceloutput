<?php
namespace mbright\ExcelOutput;

/**
 * Class SpreadSheet
 *
 * Keeps track of all the information regarding a particular spreadsheet.
 *
 * The actual data that goes into the spreadsheet must be in array form. The keys will not get used. Each entry in the
 * array represents an entry. The first entry ($data[0]) should be any desired headers. Extra metadata can get stored in
 * a array on the $meta property. Associated formatRules can be stored in an array of FormatRules objects.
 *
 * @package mbright\ExcelOutput
 */
class SpreadSheet 
{
    /**
     * Data that goes in the spreadsheet
     *
     * @var array
     */
    protected $data;

    /**
     * Extra metadata that describes the spreadsheet.
     *
     * @var array
     */
    protected $meta;

    /**
     * Array of FormatRules that apply to the spreadsheet.
     *
     * @var array
     */
    protected $formatRules;

    /**
     * @param array $data
     * @param array $meta
     * @param array $formatRules
     */
    public function __construct(array $data = [], array $meta = [], array $formatRules = [])
    {
        if (! empty($formatRules)) {
            $exceptions = array_filter(
                $formatRules,
                function ($rule) {
                    return ! $rule instanceof FormatRule;
                }
            );
            if (count($exceptions) > 0) {
                throw new \InvalidArgumentException('formatRules array can only contain FormatRule Objects');
            }
        }
        $this->data        = $data;
        $this->meta        = $meta;
        $this->formatRules = $formatRules;
    }

    /**
     * Returns spreadsheet's data.
     *
     * @return array
     */
    public function getData()
    {
        return $this->data;
    }

    /**
     * Set the spreadsheet's data.
     *
     * @param array $data
     */
    public function setData(array $data)
    {
        $this->data = $data;
    }

    /**
     * Get associated metadata.
     *
     * @return array
     */
    public function getMeta()
    {
        return $this->meta;
    }

    /**
     * Set the associated metadata.
     *
     * @param $meta
     */
    public function setMeta(array $meta)
    {
        $this->meta = $meta;
    }

    /**
     * @param FormatRule $rule
     */
    public function addRule(FormatRule $rule)
    {
        $this->formatRules[] = $rule;
    }

    /**
     * @return array
     */
    public function getFormatRules()
    {
        return $this->formatRules;
    }

    /**
     * @throws \InvalidArgumentException When array does not contain only FormatRule objects
     * @param array $formatRules contains only FormatRule objects
     */
    public function setFormatRules(array $formatRules)
    {
        $exceptions = array_filter(
                $formatRules,
                function ($rule) {
                    return ! $rule instanceof FormatRule;
                }
            );
        if (count($exceptions) > 0) {
            throw new \InvalidArgumentException('formatRules array can only contain FormatRule Objects');
        }
        $this->formatRules = $formatRules;
    }
}