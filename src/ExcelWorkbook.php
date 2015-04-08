<?php
namespace mbright\ExcelOutput;

/**
 * Class ExcelWorkbook
 *
 * Used to keep track of all data that relates to Excel Workbooks
 *
 * All SpreadSheet objects should get stored in $sheets. A Parambag is used to keep track of any extra meta information.
 * That info is stored in the $meta property.
 *
 * @package mbright\ExcelOutput
 */
class ExcelWorkbook 
{
    /**
     * Used to keep track of all SpreadSheet objects associated with the workbook.
     *
     * @var array
     */
    protected $sheets;

    /**
     * Repository for all extra meta data.
     *
     * @var array
     */
    protected $meta;

    /**
     * @param array $sheets
     * @param array $meta
     */
    public function __construct(array $sheets = null, array $meta = array())
    {
        if ($this->validateSheets($sheets)) {
            $this->sheets = $sheets;
        }
    }

    /**
     * Internal validation method for the $sheets array.
     *
     * Ensures that all objects in the given $sheets array are instnaceof SpreadSheet
     *
     * @param array $arr
     * @return bool
     */
    protected function validateSheets(array $arr)
    {
        $check = array_filter(
            $arr,
            function ($value) {
                return $value instanceof SpreadSheet;
            }
        );
        return count($check) > 0;
    }

    /**
     * Returns the $sheets array
     *
     * @return array
     */
    public function getSheets()
    {
        return $this->sheets;
    }

    /**
     * Adds a SpreadSheet object to the workbook.
     *
     * @param SpreadSheet $sheet SpreadSheet object to set
     */
    public function addSheet(SpreadSheet $sheet)
    {
        $this->sheets[] = $sheet;
    }

    /**
     * Returns the Parambag object that contains any meta data associated with the workbook.
     *
     * @return array
     */
    public function getMeta()
    {
        return $this->meta;
    }

    /**
     * Sets the given array as the workbooks metadata.
     *
     * @param array $meta
     */
    public function setMeta(array $meta)
    {
        $this->meta = $meta;
    }
}