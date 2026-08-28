<?php

declare(strict_types=1);

namespace Calculator\Calculator;

use Calculator\Support\CostMath;

/**
 * Расчёт железнодорожных ставок: типы контейнеров, COC/SOC, охрана, опасный груз.
 */
final class RailCalculator
{
    public function costForContainerType(
        string $containerType,
        array $data,
        bool $isDanger = false,
        string $ownership = 'coc',
    ): float {
        if ($isDanger) {
            return match ($containerType) {
                '20DC (<24t)', '20DC' => CostMath::positiveCeil($data['OPASNYY_20DC_24T'] ?? 0),
                '20DC (24t-28t)' => CostMath::positiveCeil($data['OPASNYY_20DC_24T_28T'] ?? 0),
                '40HC (28t)', '40HC' => CostMath::positiveCeil($data['OPASNYY_40HC_28T'] ?? 0),
                default => 0.0,
            };
        }

        return match ($containerType) {
            '20DC (<24t)', '20DC' => CostMath::positiveCeil(
                $data[$ownership === 'soc' ? 'DC20_24' : 'COC_20DC_24T'] ?? 0
            ),
            '20DC (24t-28t)' => CostMath::positiveCeil(
                $data[$ownership === 'soc' ? 'DC20_24T_28T' : 'COC_DC_24T_28T'] ?? 0
            ),
            '40HC (28t)', '40HC' => CostMath::positiveCeil(
                $data[$ownership === 'soc' ? 'HC40_28T' : 'COC_HC_28T'] ?? 0
            ),
            default => 0.0,
        };
    }

    public function securityCost(array $data, string $security, string $containerType): float
    {
        if ($security === 'no') {
            return 0.0;
        }

        $is40Hc = $containerType === '40HC (28t)' || $containerType === '40HC';
        $securityField = $is40Hc ? 'OKHRANA_40_FUT' : 'OKHRANA_20_FUT';

        if (
            ($security === '20' && !$is40Hc)
            || ($security === '40' && $is40Hc)
        ) {
            return CostMath::positiveCeil($data[$securityField] ?? 0);
        }

        // Совместимость со старым getSecurityCostForContainerType:
        // если размер охраны задан — берём поле по выбранному размеру.
        if ($security === '20' || $security === '40') {
            $field = $security === '20' ? 'OKHRANA_20_FUT' : 'OKHRANA_40_FUT';
            return CostMath::positiveCeil($data[$field] ?? 0);
        }

        return 0.0;
    }

    /**
     * @return array{20: float, 20_28: float, 40: float}
     */
    public function costsBySize(array $data, bool $isDanger = false): array
    {
        if ($isDanger) {
            return [
                '20' => CostMath::positiveCeil($data['OPASNYY_20DC_24'] ?? $data['OPASNYY_20DC_24T'] ?? 0),
                '20_28' => CostMath::positiveCeil($data['OPASNYY_DC20_24T_28T'] ?? $data['OPASNYY_20DC_24T_28T'] ?? 0),
                '40' => CostMath::positiveCeil($data['OPASNYY_HC40_28T'] ?? $data['OPASNYY_40HC_28T'] ?? 0),
            ];
        }

        return [
            '20' => CostMath::positiveCeil($data['DC20_24'] ?? 0),
            '20_28' => CostMath::positiveCeil($data['DC20_24T_28T'] ?? 0),
            '40' => CostMath::positiveCeil($data['HC40_28T'] ?? 0),
        ];
    }

    public function securityLabel(string $security): string
    {
        return match ($security) {
            '20' => '20 фут',
            '40' => '40 фут',
            default => 'Нет',
        };
    }

    /**
     * @param array{20: float|string, 20_28: float|string, 40: float|string} $baseCosts
     * @param array{20: float|string, 20_28: float|string, 40: float|string} $totalCosts
     * @param array{20: float|string, 20_28: float|string, 40: float|string}|null $dangerBaseCosts
     * @param array{20: float|string, 20_28: float|string, 40: float|string}|null $dangerTotalCosts
     * @return array<string, mixed>
     */
    public function buildResultItem(
        array $data,
        string $containerType,
        string $ownershipType,
        string $hazardType,
        string $security,
        float $profit,
        array $baseCosts,
        array $totalCosts,
        ?array $dangerBaseCosts = null,
        ?array $dangerTotalCosts = null,
        bool $showBothOwnership = false,
        bool $showBothHazard = false,
        ?array $alternativeOption = null,
    ): array {
        $item = [
            'rail_origin' => $data['POL'] ?? '',
            'rail_destination' => $data['POD'] ?? '',
            'rail_coc' => $containerType,
            'rail_container_ownership' => $ownershipType,
            'rail_agent' => $data['AGENT'] ?? '',
            'rail_hazard' => $hazardType,
            'rail_security' => $this->securityLabel($security),
            'rail_profit' => $profit,
            'cost_base_20' => $baseCosts['20'],
            'cost_base_20_28' => $baseCosts['20_28'],
            'cost_base_40' => $baseCosts['40'],
            'cost_security' => $this->securityCost($data, $security, $containerType),
            'cost_total_20' => $totalCosts['20'],
            'cost_total_20_28' => $totalCosts['20_28'],
            'cost_total_40' => $totalCosts['40'],
            'show_both_ownership' => $showBothOwnership,
            'show_both_hazard' => $showBothHazard,
        ];

        if ($dangerBaseCosts !== null && $dangerTotalCosts !== null) {
            $item['cost_base_20_danger'] = $dangerBaseCosts['20'];
            $item['cost_base_20_28_danger'] = $dangerBaseCosts['20_28'];
            $item['cost_base_40_danger'] = $dangerBaseCosts['40'];
            $item['cost_total_20_danger'] = $dangerTotalCosts['20'];
            $item['cost_total_20_28_danger'] = $dangerTotalCosts['20_28'];
            $item['cost_total_40_danger'] = $dangerTotalCosts['40'];
        }

        if ($alternativeOption !== null) {
            if ($showBothOwnership) {
                $item['soc_option'] = $alternativeOption;
            } elseif ($showBothHazard) {
                $item['normal_option'] = $alternativeOption;
            }
        }

        return $item;
    }

    /**
     * @return array<string, mixed>
     */
    public function emptyResultItem(
        array $data,
        string $containerType,
        string $ownership,
        string $hazardType,
        string $security,
    ): array {
        return [
            'rail_origin' => $data['POL'] ?? '',
            'rail_destination' => $data['POD'] ?? '',
            'rail_coc' => $containerType,
            'rail_container_ownership' => $ownership,
            'rail_agent' => $data['AGENT'] ?? '',
            'rail_hazard' => $hazardType,
            'rail_security' => $this->securityLabel($security),
            'rail_profit' => '-',
            'cost_base_normal' => '-',
            'cost_total_normal' => '-',
            'cost_base_danger' => '-',
            'cost_total_danger' => '-',
            'cost_security' => '-',
            'show_both_ownership' => false,
            'show_both_hazard_in_columns' => false,
            'empty_result' => true,
        ];
    }

    /**
     * Упрощённый расчёт одной строки (action calculateRail).
     *
     * @return array<string, mixed>
     */
    public function calculateSimpleRow(array $value, array $params): array
    {
        $cocType = (string) ($params['rail_coc'] ?? '');
        $isHazard = ($params['rail_hazard'] ?? 'no') === 'yes';
        $security = (string) ($params['rail_security'] ?? 'no');
        $profit = (float) ($params['rail_profit'] ?? 0);

        $baseCost = match ($cocType) {
            '20DC (<24t)' => CostMath::positiveCeil(
                $isHazard
                    ? ($value['OPASNYY_20DC_24'] ?? $value['OPASNYY_20DC_24T'] ?? 0)
                    : ($value['DC20_24'] ?? 0)
            ),
            '20DC (24t-28t)' => CostMath::positiveCeil(
                $isHazard
                    ? ($value['OPASNYY_DC20_24T_28T'] ?? $value['OPASNYY_20DC_24T_28T'] ?? 0)
                    : ($value['DC20_24T_28T'] ?? 0)
            ),
            '40HC (28t)' => CostMath::positiveCeil(
                $isHazard
                    ? ($value['OPASNYY_HC40_28T'] ?? $value['OPASNYY_40HC_28T'] ?? 0)
                    : ($value['HC40_28T'] ?? 0)
            ),
            default => 0.0,
        };

        $securityCost = match ($security) {
            '20' => CostMath::positiveCeil($value['OKHRANA_20_FUT'] ?? 0),
            '40' => CostMath::positiveCeil($value['OKHRANA_40_FUT'] ?? 0),
            default => 0.0,
        };

        $totalCost = CostMath::railTotal($baseCost, $securityCost, $profit);

        $containerOwnership = (string) ($params['rail_container_ownership'] ?? 'no');
        $containerType = match ($containerOwnership) {
            'coc' => 'COC',
            'soc' => 'SOC',
            default => 'Не выбрано',
        };

        return [
            'rail_origin' => $value['POL'] ?? '',
            'rail_destination' => $value['POD'] ?? '',
            'rail_coc' => $cocType,
            'rail_container_ownership' => $containerType,
            'rail_agent' => $value['AGENT'] ?? '',
            'rail_hazard' => $isHazard ? 'Да' : 'Нет',
            'rail_security' => $this->securityLabel($security),
            'cost_base' => $baseCost,
            'cost_security' => $securityCost,
            'cost_total' => $totalCost,
            'calculation_formula' => "$baseCost (базовая) + $securityCost (охрана) + $profit (прибыль) = $totalCost",
        ];
    }
}
