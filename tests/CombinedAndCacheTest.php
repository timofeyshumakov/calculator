<?php

declare(strict_types=1);

namespace Calculator\Tests;

use Calculator\Calculator\CombinedCalculator;
use Calculator\Transport\TransportCache;
use PHPUnit\Framework\TestCase;

final class CombinedAndCacheTest extends TestCase
{
    public function testCombinedRemarkAndPort(): void
    {
        $calc = new CombinedCalculator();
        $comb = [
            [
                'STANTSIYA_OTPRAVLENIYA' => 'Москва-Товарная',
                'PUNKT_OTPRAVLENIYA' => 'VVO',
                'REMARK' => 'rail note',
            ],
        ];

        $this->assertSame('VVO', $calc->findTransshipmentPort($comb, 'Москва-Товарная'));
        $this->assertSame(
            'sea note; rail note',
            $calc->combinedRemark(['REMARK' => 'sea note'], $comb, 'Москва-Товарная')
        );
    }

    public function testTransportCacheRoundtrip(): void
    {
        $dir = sys_get_temp_dir() . '/calc_cache_' . uniqid('', true);
        $cache = new TransportCache($dir, 3600);

        $this->assertNull($cache->read('transport_28'));
        $cache->write('transport_28', [['POL' => 'A']]);
        $this->assertTrue($cache->isFresh('transport_28'));
        $this->assertSame([['POL' => 'A']], $cache->read('transport_28'));

        $cache->clear(28);
        $this->assertNull($cache->read('transport_28'));

        @rmdir($dir);
    }
}
