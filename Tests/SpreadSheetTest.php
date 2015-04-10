<?php
namespace ExcelOutput\Test;

use mbright\ExcelOutput\SpreadSheet;

/**
 * Test SpreadSheetTest
 * @package ExcelOutput\Test
 */
class SpreadSheetTest extends \PHPUnit_Framework_TestCase
{
    /** @var  SpreadSheet */
    private $spreadSheet;

    /** @var array */
    private $meta = ['metadata'];

    /** @var array */
    private $formatRules;

    /** @var array */
    private $data = ['data'];

    public function setUp()
    {
        $mockRule = \Mockery::mock('mbright\ExcelOutput\FormatRule');
        $this->formatRules = array($mockRule);
        $this->spreadSheet = new SpreadSheet($this->data, $this->meta, $this->formatRules);
    }

    public function testConstructPopulated()
    {
        $this->assertEquals($this->meta, $this->spreadSheet->getMeta());
        $this->assertEquals($this->data, $this->spreadSheet->getData());
        $this->assertEquals($this->formatRules, $this->spreadSheet->getFormatRules());
    }

    public function testConstructorEmpty()
    {
        $spreadSheet = new SpreadSheet();
        $this->assertInstanceof('mbright\ExcelOutput\SpreadSheet', $spreadSheet);
        $this->assertEmpty($spreadSheet->getData());
        $this->assertEmpty($spreadSheet->getMeta());
        $this->assertEmpty($spreadSheet->getFormatRules());
    }

    public function testConstructorValidatesFormatRules()
    {
        $formatRules = array('wrong');
        $this->setExpectedException(
            'InvalidArgumentException',
            'formatRules array can only contain FormatRule Objects'
        );
        new SpreadSheet([], [], $formatRules);
    }

    public function testGetData()
    {
        $this->assertEquals($this->data, $this->spreadSheet->getData());
    }

    public function testSetData()
    {
        $data2 = ['newData'];
        $this->spreadSheet->setData($data2);
        $this->assertEquals($data2, $this->spreadSheet->getData());
    }

    public function testGetMeta()
    {
        $this->assertEquals($this->meta, $this->spreadSheet->getMeta());
    }

    public function testSetMeta()
    {
        $meta2 = ['newMeta'];
        $this->spreadSheet->setMeta($meta2);
        $this->assertEquals($meta2, $this->spreadSheet->getMeta());
    }

    public function testGetFormatRules()
    {
        $this->assertEquals($this->formatRules, $this->spreadSheet->getFormatRules());
    }

    public function testSetFormatRules()
    {
        $rule = \Mockery::mock('mbright\ExcelOutput\FormatRule');
        $rule->unique = 'uniqueValue';
        $this->spreadSheet->setFormatRules(array($rule));
        $this->assertEquals(array($rule), $this->spreadSheet->getFormatRules());
    }

    public function testSetFormatRulesException()
    {
        $formatRules = array('wrong');
        $this->setExpectedException(
            'InvalidArgumentException',
            'formatRules array can only contain FormatRule Objects'
        );
        $this->spreadSheet->setFormatRules($formatRules);
    }
}
