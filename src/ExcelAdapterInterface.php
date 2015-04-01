<?php
namespace mbright\ExcelOutput;

/**
 * Interface ExcelAdapterInterface
 *
 * The home for all boilerplate code around exporting to a given file format.
 *
 * @package mbright\ExcelOutput
 */
interface ExcelAdapterInterface 
{
    /**
     * Returns an array of all supported formats.
     *
     * @return array
     */
    public function getFormats();

    /**
     * Returns true if the given format is supported and false if not.
     *
     * @param string $format
     * @return bool
     */
    public function supportsFormat($format);

    /**
     * Handles exporting the workbook to the given format.
     *
     * Will use the data from the given workbook, and the spreadsheets it contains, to create a excel file of the given
     * format. By default, the underlying library should write to the file system, if applicable. This behavior can get
     * disabled by setting $write = false.
     *
     * @param ExcelWorkbook $workbook
     * @param string        $format
     * @param bool          $write
     * @throws \Exception   When the format is invalid.
     * @return mixed
     */
    public function exportToFormat(ExcelWorkbook $workbook, $format, $write = true);
}