<?php
namespace mbright\ExcelOutput\PHPExcel;

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
    protected $formats = [
        self::XLSX,
        self::XLS,
        self::CSV
    ];

    /** @var ExcelFormatterInterface */
    protected $formatter;

    /** @var PHPExcelFactory */
    protected $factory;

    /**
     * @param ExcelFormatterInterface $formatter
     * @param PHPExcelFactory $factory
     */
    public function __construct(ExcelFormatterInterface $formatter, PHPExcelFactory $factory)
    {
        $this->formatter = $formatter;
        $this->factory   = $factory;
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
        $phpExcel = $this->factory->newPHPExcel();
        $this->handleWorkbookMetaData($workbook, $phpExcel);
        /** @var SpreadSheet $sheet */
        foreach ($workbook->getSheets() as $sheet) {
            $phpExcel->addSheet($this->formatter->applyRules($sheet));
        }
        if ($write === true) {
            $writer   = $this->factory->newWriter($format, $phpExcel);
            $filePath = isset($workbook->getMeta()['filePath']) ? $workbook->getMeta()['filePath'] : null;
            $writer->save($filePath);
        }
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
        foreach ($workbook->getMeta() as $meta) {
            switch (true) {
                case isset($meta[self::CREATOR]):
                    $PHPExcel->getProperties()->setCreator($meta[self::CREATOR]);
                    break;
                case isset($meta[self::LAST_MODIFIED_BY]):
                    $PHPExcel->getProperties()->setLastModifiedBy($meta[self::LAST_MODIFIED_BY]);
                    break;
                case isset($meta[self::TITLE]):
                    $PHPExcel->getProperties()->setTitle($meta[self::TITLE]);
                    break;
                case isset($meta[self::SUBJECT]):
                    $PHPExcel->getProperties()->setSubject($meta[self::SUBJECT]);
                    break;
                case isset($meta[self::DESCRIPTION]):
                    $PHPExcel->getProperties()->setDescription($meta[self::DESCRIPTION]);
                    break;
                case isset($meta[self::KEYWORDS]):
                    $PHPExcel->getProperties()->setKeywords(implode(', ', $meta[self::KEYWORDS]));
                    break;
                case isset($meta[self::CATEGORIES]):
                    $PHPExcel->getProperties()->setCategory(implode(', ', $meta[self::CATEGORIES]));
            }
        }
        return $PHPExcel;
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
        return \PHPExcel_Settings::setCacheStorageMethod($method, $options);
    }
}