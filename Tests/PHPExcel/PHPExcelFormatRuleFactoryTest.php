<?php
namespace ExcelOutput\Test;

use mbright\ExcelOutput\PHPExcel\PHPExcelFormatRuleFactory;

/**
 * Test PHPExcelFormatRuleFactoryTest
 * @package ExcelOutput\Test
 */
class PHPExcelFormatRuleFactoryTest extends \PHPUnit_Framework_TestCase
{
    /** @var  PHPExcelFormatRuleFactory */
    private $factory;

    /** @var string */
    private $range = 'A:B';

    /** @var array */
    private $styleArray = array('styleRules');

    public function setUp()
    {
        $this->factory = new PHPExcelFormatRuleFactory();
    }

    public function testCreateStyleRule()
    {
        $rule = $this->factory->createStyleRule($this->range,$this->styleArray);
        $this->assertInstanceof('mbright\ExcelOutput\FormatRule', $rule);
        $this->assertEquals('style', $rule->getType());
        $this->assertEquals($this->range, $rule->getRange());
        $this->assertEquals($this->styleArray, $rule->getRules());
    }

    public function testCreateNumFormatRule()
    {
        $rule = $this->factory->createNumFormatRule($this->range, $this->styleArray);
        $this->assertInstanceof('mbright\ExcelOutput\FormatRule', $rule);
        $this->assertEquals('numFormat', $rule->getType());
        $this->assertEquals($this->range, $rule->getRange());
        $this->assertEquals($this->styleArray, $rule->getRules());
    }
}
