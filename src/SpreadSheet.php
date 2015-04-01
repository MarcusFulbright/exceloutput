<?php
namespace mbright\ExcelOutput;

/**
 * Class SpreadSheet
 *
 * Keeps track of all the information regarding a particular spreadsheet.
 *
 * The actual data that goes into the spreadsheet must be in array form. The keys will not get used. Each entry in the
 * array represents an entry. The first entry ($data[0]) should be any desired headers. Extra metadata can get stored in
 * a Parambag on the $meta property. Associated formatRules can be stored in an array of FormatRule objects.
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
     * @var \mbright\ExcelOutput\Parambag
     */
    protected $meta;

    /**
     * Array of FormatRules that apply to the spreadsheet.
     *
     * @var array
     */
    protected $formatRules;

    /**
     * @param array    $data
     * @param Parambag $meta
     * @param array    $formatRules
     */
    public function __construct(array $data = null, Parambag $meta = null, array $formatRules = null)
    {
        if ($data != null) {
            $this->data = $data;
        }

        if ($this->validateFormatRules($formatRules)) {
            $this->formatRules = $formatRules;
        }

        $this->meta = ($meta != null ? $meta : new Parambag());
    }

    /**
     * Used to validate that all array items are instances of FormatRule.
     *
     * @param array $arr
     * @return bool
     */
    protected function validateFormatRules(array $arr)
    {
        $check = array_filter(
            $arr,
            function ($value) {
                return $value instanceof FormatRule;
            }
        );
        return count($check) > 0;
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
     * @return Parambag
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
    public function setMeta(Parambag $meta)
    {
        $this->meta = $meta;
    }
}