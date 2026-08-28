<?php

declare(strict_types=1);

namespace Calculator\Support;

/**
 * Нормализация значений свойств элементов списков Bitrix24.
 */
final class BitrixValueNormalizer
{
    public static function normalize(mixed $value): string
    {
        if ($value === null || $value === '') {
            return '';
        }

        if (is_array($value)) {
            if (array_key_exists('VALUE', $value)) {
                return self::normalize($value['VALUE']);
            }
            if (array_key_exists('value', $value)) {
                return self::normalize($value['value']);
            }

            $first = reset($value);
            return $first === false ? '' : self::normalize($first);
        }

        if (is_object($value)) {
            return self::normalize((array) $value);
        }

        return is_scalar($value) ? trim((string) $value) : '';
    }

    /**
     * @param array<string, mixed> $item
     * @param array<string, string> $map oldKey => newKey
     * @return array<string, string>
     */
    public static function mapRow(array $item, array $map): array
    {
        $row = [];
        foreach ($map as $oldKey => $newKey) {
            if (!array_key_exists($oldKey, $item)) {
                continue;
            }
            $row[$newKey] = self::normalize($item[$oldKey]);
        }

        return $row;
    }

    /**
     * @param list<array<string, mixed>> $elements
     * @param array<string, string> $map
     * @return list<array<string, string>>
     */
    public static function mapElements(array $elements, array $map): array
    {
        return array_map(
            static fn(array $item): array => self::mapRow($item, $map),
            $elements
        );
    }
}
