<?php

declare(strict_types=1);

namespace Calculator\Tests;

use Calculator\Support\BitrixValueNormalizer;
use PHPUnit\Framework\TestCase;

final class BitrixValueNormalizerTest extends TestCase
{
    public function testNormalizeScalarAndNested(): void
    {
        $this->assertSame('', BitrixValueNormalizer::normalize(null));
        $this->assertSame('Shanghai', BitrixValueNormalizer::normalize('  Shanghai  '));
        $this->assertSame('Vostochny', BitrixValueNormalizer::normalize(['VALUE' => 'Vostochny']));
        $this->assertSame('A', BitrixValueNormalizer::normalize(['value' => ['VALUE' => 'A']]));
        $this->assertSame('first', BitrixValueNormalizer::normalize(['first', 'second']));
    }

    public function testMapRow(): void
    {
        $row = BitrixValueNormalizer::mapRow(
            [
                'NAME' => 'POL1',
                'PROPERTY_126' => ['VALUE' => 'POD1'],
                'IGNORED' => 'x',
            ],
            [
                'NAME' => 'POL',
                'PROPERTY_126' => 'POD',
            ]
        );

        $this->assertSame(['POL' => 'POL1', 'POD' => 'POD1'], $row);
    }
}
