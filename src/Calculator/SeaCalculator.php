<?php

declare(strict_types=1);

namespace Calculator\Calculator;

use Calculator\Support\CostMath;

/**
 * Расчёт морских ставок: COC/SOC, CAF, drop-off, опасный груз.
 */
final class SeaCalculator
{
    /**
     * @param array<string, mixed> $data
     * @return array{
     *   coc_normal: float,
     *   soc_normal: float,
     *   danger: float,
     *   coc_40_normal: float,
     *   40_danger: float,
     *   soc_40_normal: float,
     *   drop_off_20: float,
     *   drop_off_40: float
     * }
     */
    public function extractCosts(array $data, bool $is20Gp): array
    {
        $cocNormal = $is20Gp ? ($data['COC_20GP'] ?? 0) : ($data['COC_40HC'] ?? 0);
        $danger = $is20Gp ? ($data['OPASNYY_20GP'] ?? 0) : ($data['OPASNYY_40HC'] ?? 0);
        $socNormal = $is20Gp ? ($data['SOC_20GP'] ?? 0) : ($data['SOC_40HC'] ?? 0);

        return [
            'coc_normal' => CostMath::positiveCeil($cocNormal),
            'soc_normal' => CostMath::positiveCeil($socNormal),
            'danger' => CostMath::positiveCeil($danger),
            'coc_40_normal' => CostMath::positiveCeil($data['COC_40HC'] ?? 0),
            '40_danger' => CostMath::positiveCeil($data['OPASNYY_40HC'] ?? 0),
            'soc_40_normal' => CostMath::positiveCeil($data['SOC_40HC'] ?? 0),
            'drop_off_20' => CostMath::positiveCeil($data['DROP_OFF_20GP'] ?? 0),
            'drop_off_40' => CostMath::positiveCeil($data['DROP_OFF_40HC'] ?? 0),
        ];
    }

    /**
     * @param array<string, mixed> $data
     * @return array{normal: float, danger: float}
     */
    public function costsForOwnership(array $data, string $seaContainerType, string $ownershipType): array
    {
        $is20Gp = $seaContainerType === '20DC';

        if ($ownershipType === 'coc') {
            $normal = $is20Gp ? ($data['COC_20GP'] ?? 0) : ($data['COC_40HC'] ?? 0);
        } else {
            $normal = $is20Gp ? ($data['SOC_20GP'] ?? 0) : ($data['SOC_40HC'] ?? 0);
        }
        $danger = $is20Gp ? ($data['OPASNYY_20GP'] ?? 0) : ($data['OPASNYY_40HC'] ?? 0);

        return [
            'normal' => CostMath::positiveCeil($normal),
            'danger' => CostMath::positiveCeil($danger),
        ];
    }

    public function dropOffCost(array $data, string $seaContainerType): float
    {
        $field = $seaContainerType === '20DC' ? 'DROP_OFF_20GP' : 'DROP_OFF_40HC';
        return CostMath::positiveCeil($data[$field] ?? 0);
    }

    public function securityCost(array $data, string $security, string $seaContainerType): float
    {
        if ($security === 'no') {
            return 0.0;
        }

        $field = $seaContainerType === '40HC' ? 'OKHRANA_40_FUT' : 'OKHRANA_20_FUT';
        if (
            ($security === '20' && $seaContainerType !== '40HC')
            || ($security === '40' && $seaContainerType === '40HC')
        ) {
            return CostMath::positiveCeil($data[$field] ?? 0);
        }

        return 0.0;
    }

    /**
     * @param array<string, mixed> $data
     * @return array<string, mixed>
     */
    public function buildResultItem(
        array $data,
        string $containerType,
        string $ownership,
        string $hazardType,
        float $cafPercent,
        float $profit,
        float $dropOffCost,
        mixed $normalCost,
        mixed $dangerCost,
        mixed $socNormalCost = null,
    ): array {
        $isSoc = $ownership === 'SOC';

        $containerCost = null;
        if ($isSoc && $socNormalCost && $socNormalCost !== '-') {
            $containerCost = $socNormalCost;
        } elseif ($normalCost && $normalCost !== '-') {
            $containerCost = $normalCost;
        }

        $resultItem = [
            'sea_pol' => $data['POL'] ?? '',
            'sea_pod' => $data['POD'] ?? '',
            'sea_drop_off_location' => $data['DROP_OFF_LOCATION'] ?? '',
            'sea_coc' => $containerType,
            'sea_container_ownership' => $ownership,
            'sea_agent' => $data['AGENT'] ?? '',
            'sea_remark' => $data['REMARK'] ?? '',
            'sea_hazard' => $hazardType,
            'sea_caf_percent' => $cafPercent,
            'sea_profit' => $profit,
            'cost_drop_off' => $dropOffCost,
            'show_both_ownership' => false,
            'show_both_hazard_in_columns' => true,
        ];

        if ($containerCost && $containerCost !== '-') {
            $base = (float) $containerCost;
            $nettoNormal = CostMath::netto($base, $dropOffCost);
            $totalNormal = CostMath::totalWithCaf($nettoNormal, $cafPercent, $profit);

            $resultItem['cost_container_normal'] = $base;
            $resultItem['cost_netto_normal'] = $nettoNormal;
            $resultItem['cost_total_normal'] = $totalNormal;
        } else {
            $resultItem['cost_container_normal'] = '-';
            $resultItem['cost_netto_normal'] = '-';
            $resultItem['cost_total_normal'] = '-';
        }

        if ($dangerCost && $dangerCost !== '-') {
            $dangerBase = (float) $dangerCost;
            $nettoDanger = CostMath::netto($dangerBase, $dropOffCost);
            $totalDanger = CostMath::totalWithCaf($nettoDanger, $cafPercent, $profit);

            $resultItem['cost_container_danger'] = $dangerBase;
            $resultItem['cost_netto_danger'] = $nettoDanger;
            $resultItem['cost_total_danger'] = $totalDanger;
            $resultItem['show_both_hazard_in_columns'] = true;
        } else {
            $resultItem['cost_container_danger'] = '-';
            $resultItem['cost_netto_danger'] = '-';
            $resultItem['cost_total_danger'] = '-';
        }

        return $resultItem;
    }

    /**
     * @param array<string, mixed> $data
     * @return array<string, mixed>
     */
    public function emptyResultItem(
        array $data,
        string $containerType,
        string $ownership,
        string $hazardType,
        float $cafPercent = 0.0,
        float $profit = 0.0,
    ): array {
        return [
            'sea_pol' => $data['POL'] ?? '',
            'sea_pod' => $data['POD'] ?? '',
            'sea_drop_off_location' => $data['DROP_OFF_LOCATION'] ?? '',
            'sea_coc' => $containerType,
            'sea_container_ownership' => $ownership,
            'sea_agent' => $data['AGENT'] ?? '',
            'sea_remark' => $data['REMARK'] ?? '',
            'sea_hazard' => $hazardType,
            'sea_caf_percent' => $cafPercent,
            'sea_profit' => $profit,
            'cost_container_normal' => '-',
            'cost_netto_normal' => '-',
            'cost_total_normal' => '-',
            'cost_container_danger' => '-',
            'cost_netto_danger' => '-',
            'cost_total_danger' => '-',
            'cost_drop_off' => '-',
            'show_both_ownership' => false,
            'show_both_hazard_in_columns' => false,
            'empty_result' => true,
        ];
    }

    public function mapRailToSeaContainerType(string $railContainerType): string
    {
        return match ($railContainerType) {
            '20DC (<24t)', '20DC (24t-28t)', '20DC' => '20DC',
            '40HC (28t)', '40HC' => '40HC',
            default => $railContainerType,
        };
    }
}
