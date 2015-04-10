<?php
namespace ExcelOutput\PHPExcel;

use mbright\ExcelOutput\ExcelAdapterInterface;
use mbright\ExcelOutput\ExcelFormatterInterface;
use mbright\ExcelOutput\ExcelWorkbook;
use mbright\ExcelOutput\SpreadSheet;

/**
 * Class PHPExcelAdapter
 * @package ExcelOutput\PHPExcel
 */
class PHPExcelAdapter implements ExcelAdapterInterface
{
    /** @var string */
    const XLSX = 'xlsx';

    /** @var string */
    const XLS  = 'xls';

    /** @var string */
    const CSV  = 'csv';

    /** @var string */
    const CREATOR = 'creator';

    /** @var string */
    const LAST_MODIFIED_BY = 'last modified by';

    /** @var string */
    const TITLE = 'title';

    /** @var string */
    const SUBJECT = 'subject';

    /** @var string */
    const DESCRIPTION = 'description';

    /** @var string */
    const KEYWORDS = 'keywords';

    /** @var string */
    const CATEGORIES = 'categories';

    /** @var array */
    protected $formats = array(
        self::XLSX,
        self::XLS,
        self::CSV
    );

    /** @var ExcelFormatterInterface */
    protected $formatter;

    /**
     * @param ExcelFormatterInterface $formatter
     */
    public function __construct(ExcelFormatterInterface $formatter)
    {
        $this->formatter = $formatter;
    }

    /**
     * Handles exporting the workbook to the given format.
     *
     * Will use the data from the given workbook, and the spreadsheets it contains, to create a excel file of the given
     * format. By default, the underlying library should write to the file system, if applicable. This behavior can get
     * disabled by setting $write = false.
     *
     * @param ExcelWorkbook              $workbook
     * @param string                     $format
     * @param bool                       $write
     * @throws \InvalidArgumentException When the format is invalid.
     * @return mixed
     */
    public function exportToFormat(ExcelWorkbook $workbook, $format, $write = true)
    {
        if (! $this->supportsFormat($format)) {
            throw new \InvalidArgumentException('The given format is not supported');
        }
        $phpExcel = new \PHPExcel();
        $this->handleWorkbookMetaData($workbook, $phpExcel);
        /** @var SpreadSheet $sheet */
        foreach ($workbook->getSheets() as $sheet) {
            $phpExcel->addSheet($this->formatter->applyRules($sheet));
        }
        $writer = $this->getWriterForFormat($format, $phpExcel);
        $filePath = isset($workbook->getMeta()['filePath']) ? $workbook->getMeta()['filePath'] : null;
        $writer->save($workbook->$filePath);
        return $phpExcel;
    }

    /**
     * Returns a PHPExcel_Worksheet object with all the format rules applied
     *
     * @param SpreadSheet $sheet
     * @return mixed
     */
    public function formatSheet(SpreadSheet $sheet)
    {
        return $this->formatter->applyRules($sheet);
    }

    /**
     * Applies all meta data form a ExcelWorkbook to a PHPExcel object.
     *
     * @param ExcelWorkbook $workbook
     * @param \PHPExcel $PHPExcel
     * @return \PHPExcel
     */
    protected function handleWorkbookMetaData(ExcelWorkbook $workbook, \PHPExcel $PHPExcel)
    {
        $meta = $workbook->getMeta();
        switch (true) {
            case isset($meta[self::CREATOR]):
                $PHPExcel->getProperties()->setCreator($meta[self::CREATOR]);
            case isset($meta[self::LAST_MODIFIED_BY]):
                $PHPExcel->getProperties()->setLastModifiedBy($meta[self::LAST_MODIFIED_BY]);
            case isset($meta[self::TITLE]):
                $PHPExcel->getProperties()->setTitle($meta[self::TITLE]);
            case isset($meta[self::SUBJECT]):
                $PHPExcel->getProperties()->setSubject($meta[self::SUBJECT]);
            case isset($meta[self::DESCRIPTION]):
                $PHPExcel->$meta[self::DESCRIPTION];
            case isset($meta[self::KEYWORDS]):
                $PHPExcel->getProperties()->setKeywords(explode(' ', $meta[self::KEYWORDS]));
            case isset($meta[self::CATEGORES]):
                $PHPExcel->getProperties()->setCategory(explode(' ', $meta[self::CATEGORES]));
        }
        return $PHPExcel;
    }

    protected function getWriterForFormat($format, $phpExcel)
    {
        switch (true) {
            case $format === self::XLSX:
                $writer =  new \PHPExcel_Writer_Excel2007($phpExcel);
                break;
            case $format === self::XLS:
                $writer = new \PHPExcel_Writer_Excel5($phpExcel);
                break;
            case $format === self::CSV:
                $writer = new \PHPExcel_Writer_CSV($phpExcel);
                break;
            default:
                throw new \InvalidArgumentException('the given format is not supported');
        }
        return $writer;
    }

        /**
     * Returns an array of all supported formats.
     *
     * @return array
     */
    public function getFormats()
    {
        return $this->formats;
    }

    /**
     * Returns true if the given format is supported and false if not.
     *
     * @param string $format
     * @return bool
     */
    public function supportsFormat($format)
    {
        if (! in_array($format, $this->formats)) {
            return false;
        }
        return true;
    }

    /**
     * Sets the cache method using the given options, if any.
     *
     * @param $method
     * @param array $options
     * @return bool
     */
    public function setCacheStrategy($method, array $options = array())
    {
        \PHPExcel_Settings::setCacheStorageMethod($method, $options);
        return true;
    }
}