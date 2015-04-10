<?php
namespace ExcelOutput\Test;

use mbright\ExcelOutput\FormatRule;

/**
 * Test FormatRuleTest
 * @package ExcelOutput\Test
 */
class FormatRuleTest extends \PHPUnit_Framework_TestCase
{
    /** @var string */
    private $range = 'A:B';

    /** @var array */
    private $rules = [
        'font' => [
            'bold' => true
        ],
        'alignment' => [
            'horizontal' => 'HORIZONTAL_RIGHT',
            'borders' => [
                'top' => [
                    'style' => 'BORDER_THIN'
                ]
            ],
            'fill' => [
                'type' => 'FILL_GRADIENT_LINEAR',
                'rotation' => 90,
                'startcolor' => [
                    'argb' => 'FFA0A0A0'
                ],
                'endcolor' => [
                    'argb' => 'FFFFFFFF'
                ]
            ]
        ]
    ];

    /** @var string */
    private $type = 'type';

    /** @var FormatRule */
    private $formatRule;

    public function setUp()
    {
        $this->formatRule = new FormatRule($this->range, $this->rules);
    }

    public function testConstructPopulated()
    {
        $this->assertEquals($this->range, $this->formatRule->getRange());
        $this->assertEquals($this->rules, $this->formatRule->getRules());
    }

    public function testGetRange()
    {
        $this->assertEquals($this->range, $this->formatRule->getRange());
    }

    public function testSetRange()
    {
        $range = 'B:C';
        $this->formatRule->setRange($range);
        $this->assertEquals($range, $this->formatRule->getRange());
    }

    public function testGetRules()
    {
        $this->assertEquals($this->rules, $this->formatRule->getRules());
    }

    public function testSetRules()
    {
        $rules2 = array('newRUles');
        $this->formatRule->setRules($rules2);
        $this->assertEquals($rules2, $this->formatRule->getRules());
    }

    public function testSetType()
    {
        $this->formatRule->setType($this->type);
        $this->assertEquals($this->type, $this->formatRule->getType());
    }
}
