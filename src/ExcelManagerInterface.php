<?php
namespace mbright\ExcelOutput;

/**
 * Interface ExcelManagerInterface
 *
 * The glue between formatters and adapters, the one class your app needs to interact with
 *
 * @package mbright\ExcelOutput
 */
interface ExcelManagerInterface 
{
    /**
     * Ensures the injection of the required adapter and optional formatter.
     *
     * @param ExcelAdapterInterface        $adapter
     * @param null|ExcelFormatterInterface $formatter
     */
    public function __construct(ExcelAdapterInterface $adapter, ExcelFormatterInterface $formatter = null);

    /**
     * Will export the given workbook to the given format.
     *
     * The exact behavior of this method depends on the adapter
     *
     * @param ExcelWorkbook $workbook
     * @param string        $format
     * @param bool          $write default to true
     * @throws \Exception   when the format is invalid.
     * @return mixed Return depends on the adapter
     */
    public function exportToFormat(ExcelWorkbook $workbook, $format, $write = true);

    /**
     * Creates and returns a new ExcelWorkbook object.
     *
     * All items in the $sheets array should be instances of SpreadSheet. It's smart to check this in your adapter. You
     * can also validate the contents of the $meta array if desired.
     *
     * @param  array    $sheets
     * @param  array $meta
     * @return ExcelWorkbook
     */
    public function newWorkbook(array $sheets, array $meta = array());

    /**
     * Creates a new sheet from the provided data and meta-data.
     *
     * Please note that each entry in the array represents a row. In your implementation, it might be smart to check the
     * contents of the $meta array.
     *
     * @param array    $data
     * @param array $meta
     * @return SpreadSheet
     */
    public function newSheet(array $data, array $meta = array());

    /**
     * Contains all logic for applying format rules.
     *
     * Depending on the adapter and formatter, this method could behave in several ways. Be sure to document this
     * behavior in those classes. If you are seeking docs for this behavior, check the concrete implementations you are
     * using.
     *
     * @param SpreadSheet $sheet
     * @return mixed
     */
    public function formatSheet(SpreadSheet $sheet);

    /**
     * Adds a SpreadSheet to a Workbook.
     *
     * @param ExcelWorkbook $workbook
     * @param SpreadSheet   $sheet
     * @return ExcelWorkbook
     */
    public function addSheetToWorkbook(ExcelWorkbook $workbook, SpreadSheet $sheet);

    /**
     * Returns true if the adapter supports the given format, false if it does not.
     *
     * @param $format
     * @return bool
     */
    public function supportsFormat($format);
}