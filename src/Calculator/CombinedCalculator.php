<?php

declare(strict_types=1);

namespace Calculator\Calculator;

/**
 * Вспомогательная логика комбинированных перевозок (море + ЖД).
 */
final class CombinedCalculator
{
    /**
     * @param list<array<string, mixed>> $combPerevozki
     */
    public function findTransshipmentPort(array $combPerevozki, string $railStartStation): string
    {
        if ($combPerevozki === [] || $railStartStation === '') {
            return '';
        }

        foreach ($combPerevozki as $row) {
            if (trim((string) ($row['STANTSIYA_OTPRAVLENIYA'] ?? '')) === trim($railStartStation)) {
                return (string) ($row['PUNKT_OTPRAVLENIYA'] ?? '');
            }
        }

        return '';
    }

    /**
     * @param array<string, mixed> $seaValue
     * @param list<array<string, mixed>> $combPerevozki
     */
    public function combinedRemark(array $seaValue, array $combPerevozki, string $railStartStation): string
    {
        $remarks = [];

        if (!empty($seaValue['REMARK'])) {
            $remarks[] = trim((string) $seaValue['REMARK']);
        }

        foreach ($combPerevozki as $row) {
            if (trim((string) ($row['STANTSIYA_OTPRAVLENIYA'] ?? '')) !== trim($railStartStation)) {
                continue;
            }
            if (!empty($row['REMARK'])) {
                $remarks[] = trim((string) $row['REMARK']);
            }
            break;
        }

        return implode('; ', $remarks);
    }

    /**
     * @param array<string, mixed> $seaData
     * @return array<string, mixed>
     */
    public function emptyResultItem(
        array $seaData,
        string $containerType,
        string $ownership,
        string $hazardType,
        string $railOrigin = '',
        string $railDestination = '',
    ): array {
        return [
            'sea_pol' => $seaData['POL'] ?? '',
            'sea_pod' => $seaData['POD'] ?? '',
            'rail_origin' => $railOrigin,
            'rail_destination' => $railDestination,
            'container_type' => $containerType,
            'container_ownership' => $ownership,
            'hazard' => $hazardType,
            'cost_total' => '-',
            'empty_result' => true,
        ];
    }
}
