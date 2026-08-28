<?php

declare(strict_types=1);

namespace Calculator\Tests;

use Calculator\Calculator\RailCalculator;
use PHPUnit\Framework\TestCase;

final class RailCalculatorTest extends TestCase
{
    private RailCalculator $calc;

    protected function setUp(): void
    {
        $this->calc = new RailCalculator();
    }

    public function testCostForContainerTypeCocAndSoc(): void
    {
        $data = [
            'COC_20DC_24T' => 1000,
            'DC20_24' => 800,
            'OPASNYY_20DC_24T' => 1500,
            'OKHRANA_20_FUT' => 100,
            'OKHRANA_40_FUT' => 200,
        ];

        $this->assertSame(1000.0, $this->calc->costForContainerType('20DC (<24t)', $data, false, 'coc'));
        $this->assertSame(800.0, $this->calc->costForContainerType('20DC (<24t)', $data, false, 'soc'));
        $this->assertSame(1500.0, $this->calc->costForContainerType('20DC (<24t)', $data, true, 'coc'));
    }

    public function testSecurityCostMatchesContainerSize(): void
    {
        $data = [
            'OKHRANA_20_FUT' => 111,
            'OKHRANA_40_FUT' => 222,
        ];

        $this->assertSame(0.0, $this->calc->securityCost($data, 'no', '20DC (<24t)'));
        $this->assertSame(111.0, $this->calc->securityCost($data, '20', '20DC (<24t)'));
        $this->assertSame(222.0, $this->calc->securityCost($data, '40', '40HC (28t)'));
    }

    public function testCalculateSimpleRow(): void
    {
        $row = $this->calc->calculateSimpleRow(
            [
                'POL' => 'A',
                'POD' => 'B',
                'AGENT' => 'X',
                'DC20_24' => 1000.2,
                'OKHRANA_20_FUT' => 50.5,
            ],
            [
                'rail_coc' => '20DC (<24t)',
                'rail_hazard' => 'no',
                'rail_security' => '20',
                'rail_profit' => 10,
                'rail_container_ownership' => 'coc',
            ]
        );

        $this->assertSame(1001.0, $row['cost_base']);
        $this->assertSame(51.0, $row['cost_security']);
        $this->assertSame(1062.0, $row['cost_total']);
        $this->assertSame('COC', $row['rail_container_ownership']);
    }
}
