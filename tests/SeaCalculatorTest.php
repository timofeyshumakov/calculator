<?php

declare(strict_types=1);

namespace Calculator\Tests;

use Calculator\Calculator\SeaCalculator;
use PHPUnit\Framework\TestCase;

final class SeaCalculatorTest extends TestCase
{
    private SeaCalculator $calc;

    protected function setUp(): void
    {
        $this->calc = new SeaCalculator();
    }

    public function testExtractCostsFor20Dc(): void
    {
        $costs = $this->calc->extractCosts([
            'COC_20GP' => '100.2',
            'SOC_20GP' => '90',
            'OPASNYY_20GP' => '150.5',
            'COC_40HC' => '200',
            'SOC_40HC' => '180',
            'OPASNYY_40HC' => '250',
            'DROP_OFF_20GP' => '10.1',
            'DROP_OFF_40HC' => '20',
        ], true);

        $this->assertSame(101.0, $costs['coc_normal']);
        $this->assertSame(90.0, $costs['soc_normal']);
        $this->assertSame(151.0, $costs['danger']);
        $this->assertSame(11.0, $costs['drop_off_20']);
    }

    public function testBuildResultItemAppliesCafAndProfit(): void
    {
        $item = $this->calc->buildResultItem(
            [
                'POL' => 'SHA',
                'POD' => 'VVO',
                'DROP_OFF_LOCATION' => 'VVO',
                'AGENT' => 'Agent',
                'REMARK' => 'ok',
            ],
            '20DC',
            'COC',
            'Нет',
            10.0,
            100.0,
            50.0,
            1000,
            1200
        );

        $this->assertSame(1000.0, $item['cost_container_normal']);
        $this->assertSame(1050.0, $item['cost_netto_normal']);
        $this->assertSame(1255.0, $item['cost_total_normal']);
        $this->assertSame(1200.0, $item['cost_container_danger']);
        $this->assertSame(1250.0, $item['cost_netto_danger']);
        // Исправлено: раньше в legacy писался dangerCost вместо totalDanger
        $this->assertSame(1475.0, $item['cost_total_danger']);
    }

    public function testMapRailToSeaContainerType(): void
    {
        $this->assertSame('20DC', $this->calc->mapRailToSeaContainerType('20DC (24t-28t)'));
        $this->assertSame('40HC', $this->calc->mapRailToSeaContainerType('40HC (28t)'));
    }
}
