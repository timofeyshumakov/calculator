<?php

declare(strict_types=1);

namespace Calculator\Tests;

use Calculator\Support\CostMath;
use PHPUnit\Framework\TestCase;

final class CostMathTest extends TestCase
{
    public function testPositiveCeil(): void
    {
        $this->assertSame(0.0, CostMath::positiveCeil(null));
        $this->assertSame(0.0, CostMath::positiveCeil(''));
        $this->assertSame(0.0, CostMath::positiveCeil(0));
        $this->assertSame(11.0, CostMath::positiveCeil(10.1));
        $this->assertSame(10.0, CostMath::positiveCeil(10));
    }

    public function testNettoAndTotalWithCaf(): void
    {
        $netto = CostMath::netto(1000, 50);
        $this->assertSame(1050.0, $netto);

        $total = CostMath::totalWithCaf($netto, 10, 100);
        // ceil(1050 * 1.1 + 100) = ceil(1255) = 1255
        $this->assertSame(1255.0, $total);
    }

    public function testRailTotal(): void
    {
        $this->assertSame(1350.0, CostMath::railTotal(1000, 250, 100));
    }

    public function testHasPositiveCost(): void
    {
        $this->assertFalse(CostMath::hasPositiveCost('-'));
        $this->assertFalse(CostMath::hasPositiveCost(0));
        $this->assertTrue(CostMath::hasPositiveCost(1));
    }
}
