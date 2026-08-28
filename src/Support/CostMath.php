<?php

declare(strict_types=1);

namespace Calculator\Support;

/**
 * Общие формулы округления и надбавок калькулятора.
 */
final class CostMath
{
    public static function positiveCeil(mixed $value): float
    {
        if ($value === null || $value === '' || $value === '-') {
            return 0.0;
        }

        $number = (float) $value;
        return $number > 0 ? (float) ceil($number) : 0.0;
    }

    public static function hasPositiveCost(mixed $value): bool
    {
        return self::positiveCeil($value) > 0;
    }

    /**
     * Нетто = базовая ставка + drop-off (с ceil).
     */
    public static function netto(float $base, float $dropOff = 0.0): float
    {
        return (float) ceil($base + $dropOff);
    }

    /**
     * Итог = нетто * (1 + CAF%) + прибыль (с ceil).
     */
    public static function totalWithCaf(float $netto, float $cafPercent, float $profit = 0.0): float
    {
        return (float) ceil($netto * (1 + $cafPercent / 100) + $profit);
    }

    /**
     * Ж/Д итог = база + охрана + прибыль.
     */
    public static function railTotal(float $base, float $security, float $profit = 0.0): float
    {
        return (float) ceil($base + $security + $profit);
    }

    public static function dashIfZero(float $value): float|string
    {
        return $value > 0 ? $value : '-';
    }
}
