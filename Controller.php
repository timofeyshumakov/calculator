<?php

// Подключаем PHP-SDK CRest для работы с REST API
require_once __DIR__ . '/crestV136/crest.php';
/*
// Подключаем PhpSpreadsheet напрямую (без Composer)
require_once __DIR__ . '/vendor/phpoffice/phpspreadsheet/src/PhpSpreadsheet/Autoloader.php';
\PhpOffice\PhpSpreadsheet\Autoloader::register();

// Подключаем библиотеку разбора Excel
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
*/
/**
 * Контроллер для обработки запросов калькулятора стоимости перевозок.
 * Обрабатывает два экшена: index (отображает форму) и install (устанавливает локальное приложение).
 */
class TransportationCalculatorController
{
    // Константы ID инфоблоков
    public const IBLOCK_RAIL_TRANSPORTATION = 30;
    public const IBLOCK_SEA_TRANSPORTATION = 28;
    public const IBLOCK_COMBINED_TRANSPORTATION = 32;

    // маппинг полей жд перевозок
    public const ZHD_TRANSPORT_MAP = [
        'NAME' => 'POL',
        'PROPERTY_142' => 'POD',
        'PROPERTY_166' => 'DC20_24',
        'PROPERTY_170' => 'DC20_24T_28T',
        'PROPERTY_174' => 'HC40_28T',
        'PROPERTY_178' => 'OKHRANA_20_FUT',
        'PROPERTY_180' => 'OKHRANA_40_FUT',
        'PROPERTY_196' => 'AGENT',
        'PROPERTY_212' => 'COC_20DC_24T',
        'PROPERTY_214' => 'COC_DC_24T_28T',
        'PROPERTY_216' => 'COC_HC_28T',
        'PROPERTY_168' => 'OPASNYY_20DC_24T',
        'PROPERTY_172' => 'OPASNYY_20DC_24T_28T',
        'PROPERTY_176' => 'OPASNYY_40HC_28T',
    ];

    // маппинг полей морских перевозок
    public const SEA_TRANSPORT_MAP = [
        'NAME' => 'POL',
        'PROPERTY_126' => 'POD',
        'PROPERTY_162' => 'COC_20GP',
        'PROPERTY_164' => 'COC_40HC',
        'PROPERTY_132' => 'DROP_OFF_LOCATION',
        'PROPERTY_134' => 'DROP_OFF_20GP',
        'PROPERTY_136' => 'DROP_OFF_40HC',
        'PROPERTY_138' => 'CAF_KONVERT',
        'PROPERTY_140' => 'REMARK',
        'PROPERTY_192' => 'AGENT',
        'PROPERTY_202' => 'SOC_20GP',
        'PROPERTY_200' => 'SOC_40HC',
        'PROPERTY_204' => 'OKHRANA_20_FUT',
        'PROPERTY_206' => 'OKHRANA_40_FUT',
        'PROPERTY_208' => 'OPASNYY_20GP',
        'PROPERTY_210' => 'OPASNYY_40HC',
    ];

    // маппинг комбинированных перевозок
    public const COMB_TRANSPORT_MAP = [
        'NAME' => 'POL',
        'PROPERTY_182' => 'PUNKT_OTPRAVLENIYA',
        'PROPERTY_184' => 'STANTSIYA_OTPRAVLENIYA',
        'PROPERTY_186' => 'PUNKT_NAZNACHENIYA',
        'PROPERTY_188' => 'STANTSIYA_NAZNACHENIYA',
        'PROPERTY_190' => 'REMARK',
    ];

    public const TRANSSHIPMENT_PORT_AGENT_MAP = [
        'ВСК' => 'Восточная Стивидорная Компания',
    ];

    /**
     * Экшен index: отрисовывает форму расчета.
     */
    public function index()
    {
        // данные жд перевозок
        $zhdPerevozki  = self::fetchTransportData(self::IBLOCK_RAIL_TRANSPORTATION, self::ZHD_TRANSPORT_MAP);
        // данные морских перевозок
        $seaPerevozki  = self::fetchTransportData(self::IBLOCK_SEA_TRANSPORTATION, self::SEA_TRANSPORT_MAP);
        // данные комбинированных перевозок
        $combPerevozki = self::fetchTransportData(self::IBLOCK_COMBINED_TRANSPORTATION, self::COMB_TRANSPORT_MAP);

        // Подключаем файл с формой
        $formFile = __DIR__ . '/Forms.php';
        if (file_exists($formFile)) {
            include $formFile;
        } else {
            header('HTTP/1.0 500 Internal Server Error');
            echo 'Ошибка: файл Forms.php не найден.';
        }
    }

    /**
     * Получаем морские маршруты для отображения
     *
     * @return [type]
     */
    
public function getSeaPerevozki() {
    header('Content-Type: application/json; charset=utf-8');
    $result = [];
    $params = $_POST;
    
    try {
        // Получаем данные морских перевозок
        $seaPerevozki = self::fetchTransportData(
            self::IBLOCK_SEA_TRANSPORTATION, 
            self::SEA_TRANSPORT_MAP,
            [
                '=NAME' => $params['sea_pol'] ?? '',
                'PROPERTY_126' => $params['sea_pod'] ?? '',
                'PROPERTY_132' => $params['sea_drop_off_location'] ?? '',
            ]
        );
        
        // Если не нашли по точному совпадению, ищем только по POL и POD
        if (empty($seaPerevozki) && !empty($params['sea_pol']) && !empty($params['sea_pod'])) {
            $seaPerevozki = self::fetchTransportData(
                self::IBLOCK_SEA_TRANSPORTATION, 
                self::SEA_TRANSPORT_MAP,
                [
                    '=NAME' => $params['sea_pol'],
                    'PROPERTY_126' => $params['sea_pod'],
                ]
            );
        }

        // Если не нашли данные вообще
        if (empty($seaPerevozki)) {
            echo json_encode([
                'error' => true,
                'message' => 'Не найдены данные для указанных параметров'
            ], JSON_UNESCAPED_UNICODE);
            return;
        }
        
        // Параметры расчета
        $isHazard = ($params['sea_hazard'] ?? 'no') === 'yes';
        $profit = floatval($params['sea_profit'] ?? 0);
        $containerOwnership = $params['sea_container_ownership'] ?? 'no';
        $containerType = $params['sea_coc'] ?? '';

        // Определяем типы контейнеров для отображения
        $containerTypesToShow = empty($containerType) ? ['20DC', '40HC'] : [$containerType];
        
        // Флаг для раздельного отображения COC/SOC
        $showBothOwnership = ($containerOwnership === 'no');
        
        // Обрабатываем каждую найденную запись
        foreach ($seaPerevozki as $value) {
            // Используем CAF из данных, если не передан в запросе
            $actualCafPercent = floatval($value["CAF_KONVERT"]);

            foreach ($containerTypesToShow as $cType) {
                $is20GP = ($cType === '20DC');
                
                // Получаем стоимости
                $costs = $this->getSeaCosts($value, $is20GP);
                
                $dropOffCost = $is20GP ? $costs['drop_off_20'] : $costs['drop_off_40'];

                // Если показываем оба варианта собственности
                if ($showBothOwnership) {
                    // Проверяем наличие цены для COC (обычный и опасный)
                    $hasCOCNormalCost = !empty($costs['coc_normal']) && floatval($costs['coc_normal']) > 0;
                    $hasDangerCost = !empty($costs['danger']) && floatval($costs['danger']) > 0;
                    $hasSOCNormalCost = !empty($costs['soc_normal']) && floatval($costs['soc_normal']) > 0;

                    if ($hasCOCNormalCost) {
                        $cocResult = $this->createSeaResultItem(
                            $value, $cType, 'COC', 'Оба варианта', 
                            $actualCafPercent, $profit, $dropOffCost,
                            $hasCOCNormalCost ? $costs['coc_normal'] : '-', 
                            $hasDangerCost ? $costs['danger'] : '-'
                        );

                        if (!empty($cocResult)) {
                            $result[] = $cocResult;
                        }
                    }

                    if ($hasSOCNormalCost) {
                        $socResult = $this->createSeaResultItem(
                            $value, $cType, 'SOC', 'Оба варианта',
                            $actualCafPercent, $profit, $dropOffCost,
                            $hasSOCNormalCost ? $costs['soc_normal'] : '-', 
                            $hasDangerCost ? $costs['danger'] : '-'
                        );
                        
                        if (!empty($socResult)) {
                            $result[] = $socResult;
                        }
                    }
                } else {
                    // Один вариант собственности
                    $selectedType = $containerOwnership === 'coc' ? 'COC' : 'SOC';
                    
                    // Получаем стоимость контейнера
                    $containerCost = 0;
                    $hasContainerCost = false;
                    $dangerCost = 0;
                    $hasDangerCost = false;
                    
                    if ($containerOwnership === 'coc') {
                        $containerCost = $is20GP ? $costs['coc_normal'] : $costs['coc_40_normal'];
                        $hasContainerCost = !empty($containerCost) && floatval($containerCost) > 0;
                        $dangerCost = $is20GP ? $costs['danger'] : $costs['40_danger'];
                        $hasDangerCost = !empty($dangerCost) && floatval($dangerCost) > 0;
                    } else {
                        $containerCost = $is20GP ? $costs['soc_normal'] : $costs['soc_40_normal'];
                        $hasContainerCost = !empty($containerCost) && floatval($containerCost) > 0;
                        $dangerCost = $is20GP ? $costs['danger'] : $costs['40_danger'];
                        $hasDangerCost = !empty($dangerCost) && floatval($dangerCost) > 0;
                    }
                    
                    if (!$isHazard) {
                        // Обычный груз - показываем оба варианта опасности ТОЛЬКО если есть хотя бы одна стоимость
                        if ($hasContainerCost) {
                            $normalResult = $this->createSeaResultItem(
                                $value, $cType, $selectedType, 'Нет',
                                $actualCafPercent, $profit, $dropOffCost,
                                $hasContainerCost ? $containerCost : '-', 
                                $hasDangerCost ? $dangerCost : '-'
                            );
                            
                            if (!empty($normalResult)) {
                                $result[] = $normalResult;
                            }
                        }
                    } else {
                        // Опасный груз - показываем только опасный ТОЛЬКО если есть стоимость
                        if ($hasDangerCost) {
                            $dangerResult = $this->createSeaResultItem(
                                $value, $cType, $selectedType, 'Да',
                                $actualCafPercent, $profit, $dropOffCost,
                                null,
                                $dangerCost
                            );
                            
                            if (!empty($dangerResult)) {
                                $result[] = $dangerResult;
                            }
                        }
                    }
                }
            }
        }
        
        // Если после фильтрации нет данных
        if (empty($result)) {
            $result = [
                'error' => true,
                'message' => 'Не найдены данные с указанными ценами для выбранных параметров'
            ];
        }

    } catch (\Exception $e) {
        file_put_contents(__DIR__ . '/error.log', date('c') . " (getSeaPerevozki) " . $e->getMessage() . "\n", FILE_APPEND|LOCK_EX);
        $result = [
            'error' => true,
            'message' => 'Ошибка при расчете: ' . $e->getMessage()
        ];
    }
    
    echo json_encode($result, JSON_UNESCAPED_UNICODE);
    return json_encode($result, JSON_UNESCAPED_UNICODE);
}

/**
 * Получает стоимости для морской перевозки
 */
private function getSeaCosts($data, $is20GP): array {
    
    $coc_normal = $is20GP ? ($data['COC_20GP'] ?? 0) : ($data['COC_40HC'] ?? 0);
    $danger = $is20GP ? ($data['OPASNYY_20GP'] ?? 0) : ($data['OPASNYY_40HC'] ?? 0);
    $soc_normal = $is20GP ? ($data['SOC_20GP'] ?? 0) : ($data['SOC_40HC'] ?? 0);

    return [
        'coc_normal' => !empty($coc_normal) && floatval($coc_normal) > 0 ? ceil(floatval($coc_normal)) : 0,
        'soc_normal' => !empty($soc_normal) && floatval($soc_normal) > 0 ? ceil(floatval($soc_normal)) : 0,
        'danger' => !empty($danger) && floatval($danger) > 0 ? ceil(floatval($danger)) : 0,
        'coc_40_normal' => !empty($data['COC_40HC']) && floatval($data['COC_40HC']) > 0 ? ceil(floatval($data['COC_40HC'])) : 0,
        '40_danger' => !empty($data['OPASNYY_40HC']) && floatval($data['OPASNYY_40HC']) > 0 ? ceil(floatval($data['OPASNYY_40HC'])) : 0,
        'soc_40_normal' => !empty($data['SOC_40HC']) && floatval($data['SOC_40HC']) > 0 ? ceil(floatval($data['SOC_40HC'])) : 0,
        'drop_off_20' => !empty($data['DROP_OFF_20GP']) && floatval($data['DROP_OFF_20GP']) > 0 ? ceil(floatval($data['DROP_OFF_20GP'])) : 0,
        'drop_off_40' => !empty($data['DROP_OFF_40HC']) && floatval($data['DROP_OFF_40HC']) > 0 ? ceil(floatval($data['DROP_OFF_40HC'])) : 0
    ];
}
/**
 * Создает элемент результата для морских перевозок
 */
private function createSeaResultItem($data, $containerType, $ownership, $hazardType, 
                                     $cafPercent, $profit, $dropOffCost, 
                                     $normalCost, $dangerCost, $socNormalCost = null): array {
    
    $isSOC = ($ownership === 'SOC');
    
    // Определяем базовую стоимость контейнера
    $containerCost = null;
    if ($isSOC && $socNormalCost && $socNormalCost !== '-') {
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
        'show_both_hazard_in_columns' => true
    ];
    
    // Расчет для обычного груза если есть стоимость
    if ($containerCost && $containerCost !== '-') {
        $nettoNormal = ceil($containerCost + $dropOffCost);
        $totalNormal = ceil($nettoNormal * (1 + $cafPercent / 100) + $profit);
        
        $resultItem['cost_container_normal'] = $containerCost;
        $resultItem['cost_netto_normal'] = $nettoNormal;
        $resultItem['cost_total_normal'] = $totalNormal;
    } else {
        $resultItem['cost_container_normal'] = '-';
        $resultItem['cost_netto_normal'] = '-';
        $resultItem['cost_total_normal'] = '-';
    }
    
    // Расчет для опасного груза если есть стоимость
    if ($dangerCost && $dangerCost !== '-') {
        $nettoDanger = ceil($dangerCost + $dropOffCost);
        $totalDanger = ceil($nettoDanger * (1 + $cafPercent / 100) + $profit);
        
        $resultItem['cost_container_danger'] = $dangerCost;
        $resultItem['cost_netto_danger'] = $nettoDanger;
        $resultItem['cost_total_danger'] = $dangerCost;
        $resultItem['show_both_hazard_in_columns'] = true;
    } else {
        $resultItem['cost_container_danger'] = '-';
        $resultItem['cost_netto_danger'] = '-';
        $resultItem['cost_total_danger'] = '-';
    }
    
    return $resultItem;
}
/**
 * Создает пустой элемент результата для морских перевозок (прочерк)
 */
private function createEmptySeaResultItem($data, $containerType, $ownership, $hazardType, $cafPercent = 0, $profit = 0): array {
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
        'empty_result' => true
    ];
}
/**
 * Получаем ж/д маршруты для отображения
 *
 * @return [type]
 */
public function getRailPerevozki() {
    header('Content-Type: application/json; charset=utf-8');
    $result = [];
    $params = $_POST;
    
    try {
        // Получаем данные ж/д перевозок с фильтрацией
        $zhdPerevozki = self::fetchTransportData(
            self::IBLOCK_RAIL_TRANSPORTATION, 
            self::ZHD_TRANSPORT_MAP,
            [
                '=NAME' => $params['rail_origin'] ?? '',
                '=PROPERTY_142' => $params['rail_destination'] ?? '',
            ]
        );
        
        // Если не нашли данные
        if (empty($zhdPerevozki)) {
            echo json_encode([
                'error' => true,
                'message' => 'Не найдены данные для указанных параметров'
            ], JSON_UNESCAPED_UNICODE);
            return;
        }
        
        // Параметры расчета
        $cocType = $params['rail_coc'] ?? '';
        $isHazard = ($params['rail_hazard'] ?? 'no') === 'yes';
        $security = $params['rail_security'] ?? 'no';
        $profit = floatval($params['rail_profit'] ?? 0);
        $containerOwnership = $params['rail_container_ownership'] ?? 'no';
        
        // Обрабатываем каждую найденную запись
        foreach ($zhdPerevozki as $value) {
            // Получаем стоимость охраны для выбранного типа контейнера
            $securityCost = $this->getSecurityCostForContainerType($value, $security, $cocType);
            
            // Если собственность контейнера не выбрана ('no') - показываем оба варианта в ОТДЕЛЬНЫХ РЯДАХ
            if ($containerOwnership === 'no') {
                // Проверяем стоимость для COC
                $normalCostCOC = $this->getRailCostForContainerType($cocType, $value, false, 'coc');
                $dangerCostCOC = $this->getRailCostForContainerType($cocType, $value, true, 'coc');
                $hasCostCOC = $normalCostCOC > 0;
                
                // Добавляем ряд для COC ТОЛЬКО если есть хотя бы одна стоимость
                if ($hasCostCOC) {
                    $resultItemCOC = [
                        'rail_origin' => $value['POL'] ?? '',
                        'rail_destination' => $value['POD'] ?? '',
                        'rail_coc' => $cocType,
                        'rail_container_ownership' => 'COC',
                        'rail_agent' => $value['AGENT'] ?? '',
                        'rail_hazard' => 'Оба варианта',
                        'rail_security' => $security === 'no' ? 'Нет' : ($security === '20' ? '20 фут' : '40 фут'),
                        'rail_profit' => $profit,
                        
                        // Обычный груз
                        'cost_base_normal' => $normalCostCOC > 0 ? $normalCostCOC : '-',
                        'cost_total_normal' => $normalCostCOC > 0 ? ceil($normalCostCOC + $securityCost + $profit) : '-',
                        
                        // Опасный груз
                        'cost_base_danger' => $dangerCostCOC > 0 ? $dangerCostCOC : '-',
                        'cost_total_danger' => $dangerCostCOC > 0 ? ceil($dangerCostCOC) : '-',
                        
                        // Общие поля
                        'cost_security' => $securityCost,
                        'show_both_ownership' => false,
                        'show_both_hazard_in_columns' => true
                    ];
                    $result[] = $resultItemCOC;
                }
                
                // Проверяем стоимость для SOC
                $normalCostSOC = $this->getRailCostForContainerType($cocType, $value, false, 'soc');
                $dangerCostSOC = $this->getRailCostForContainerType($cocType, $value, true, 'soc');
                $hasCostSOC = $normalCostSOC > 0;
                
                // Добавляем ряд для SOC ТОЛЬКО если есть хотя бы одна стоимость
                if ($hasCostSOC) {
                    $resultItemSOC = [
                        'rail_origin' => $value['POL'] ?? '',
                        'rail_destination' => $value['POD'] ?? '',
                        'rail_coc' => $cocType,
                        'rail_container_ownership' => 'SOC',
                        'rail_agent' => $value['AGENT'] ?? '',
                        'rail_hazard' => 'Оба варианта',
                        'rail_security' => $security === 'no' ? 'Нет' : ($security === '20' ? '20 фут' : '40 фут'),
                        'rail_profit' => $profit,
                        
                        // Обычный груз
                        'cost_base_normal' => $normalCostSOC > 0 ? $normalCostSOC : '-',
                        'cost_total_normal' => $normalCostSOC > 0 ? ceil($normalCostSOC + $securityCost + $profit) : '-',
                        
                        // Опасный груз
                        'cost_base_danger' => $dangerCostSOC > 0 ? $dangerCostSOC : '-',
                        'cost_total_danger' => $dangerCostSOC > 0 ? ceil($dangerCostSOC) : '-',
                        
                        // Общие поля
                        'cost_security' => $securityCost,
                        'show_both_ownership' => false,
                        'show_both_hazard_in_columns' => true
                    ];
                    $result[] = $resultItemSOC;
                }
                
            } else {
                // Если выбрана конкретная собственность контейнера
                $displayContainerType = $containerOwnership === 'coc' ? 'COC' : 'SOC';
                $ownershipType = $containerOwnership === 'coc' ? 'coc' : 'soc';
                
                // Получаем стоимость обычного груза для выбранного типа собственности
                $normalCost = $this->getRailCostForContainerType($cocType, $value, false, $ownershipType);
                $dangerCost = $this->getRailCostForContainerType($cocType, $value, true, $ownershipType);

                // Проверяем наличие хотя бы одной стоимости
                $hasCost = $normalCost > 0;
                
                if ($hasCost) {
                    if (!$isHazard) {
                        // Пользователь выбрал обычный груз - показываем оба варианта ТОЛЬКО если есть хотя бы одна стоимость
                        $resultItem = [
                            'rail_origin' => $value['POL'] ?? '',
                            'rail_destination' => $value['POD'] ?? '',
                            'rail_coc' => $cocType,
                            'rail_container_ownership' => $displayContainerType,
                            'rail_agent' => $value['AGENT'] ?? '',
                            'rail_hazard' => 'Нет',
                            'rail_security' => $security === 'no' ? 'Нет' : ($security === '20' ? '20 фут' : '40 фут'),
                            'rail_profit' => $profit,
                            
                            // Обычный груз
                            'cost_base_normal' => $normalCost > 0 ? $normalCost : '-',
                            'cost_total_normal' => $normalCost > 0 ? ceil($normalCost + $securityCost + $profit) : '-',
                            
                            // Опасный груз
                            'cost_base_danger' => $dangerCost > 0 ? $dangerCost : '-',
                            'cost_total_danger' => $dangerCost > 0 ? ceil($dangerCost) : '-',
                            
                            // Общие поля
                            'cost_security' => $securityCost,
                            'show_both_ownership' => false,
                            'show_both_hazard_in_columns' => true
                        ];

                        $result[] = $resultItem;
                        
                    } else {
                        // Пользователь выбрал опасный груз - показываем только опасный ТОЛЬКО если есть стоимость
                        if ($dangerCost > 0) {
                            $resultItem = [
                                'rail_origin' => $value['POL'] ?? '',
                                'rail_destination' => $value['POD'] ?? '',
                                'rail_coc' => $cocType,
                                'rail_container_ownership' => $displayContainerType,
                                'rail_agent' => $value['AGENT'] ?? '',
                                'rail_hazard' => 'Да',
                                'rail_security' => $security === 'no' ? 'Нет' : ($security === '20' ? '20 фут' : '40 фут'),
                                'rail_profit' => $profit,
                                
                                // Только опасный груз
                                'cost_base' => $dangerCost,
                                'cost_total' => ceil($dangerCost + $securityCost + $profit),
                                
                                // Общие поля
                                'cost_security' => $securityCost,
                                'show_both_ownership' => false,
                                'show_both_hazard_in_columns' => false
                            ];
                            
                            // Формула расчета
                            $resultItem['calculation_formula'] = "{$dangerCost} (базовая) + {$securityCost} (охрана) + {$profit} (прибыль) = " . ceil($dangerCost + $securityCost + $profit) . " ₽";
                            
                            $result[] = $resultItem;
                        }
                    }
                }
            }
        }
        
    } catch (\Exception $e) {
        file_put_contents(__DIR__ . '/error.log', date('c') . " (getRailPerevozki) " . $e->getMessage() . "\n", FILE_APPEND|LOCK_EX);
        $result = [
            'error' => true,
            'message' => 'Ошибка при получении данных: ' . $e->getMessage()
        ];
    }
    
    echo json_encode($result, JSON_UNESCAPED_UNICODE);
    return json_encode($result, JSON_UNESCAPED_UNICODE);
}

/**
 * Создает пустой элемент результата для ж/д перевозок (прочерк)
 */
private function createEmptyRailResultItem($data, $containerType, $ownership, $hazardType, $security): array {
    $securityText = $security === 'no' ? 'Нет' : ($security === '20' ? '20 фут' : '40 фут');
    
    return [
        'rail_origin' => $data['POL'] ?? '',
        'rail_destination' => $data['POD'] ?? '',
        'rail_coc' => $containerType,
        'rail_container_ownership' => $ownership,
        'rail_agent' => $data['AGENT'] ?? '',
        'rail_hazard' => $hazardType,
        'rail_security' => $securityText,
        'rail_profit' => '-',
        'cost_base_normal' => '-',
        'cost_total_normal' => '-',
        'cost_base_danger' => '-',
        'cost_total_danger' => '-',
        'cost_security' => '-',
        'show_both_ownership' => false,
        'show_both_hazard_in_columns' => false,
        'empty_result' => true
    ];
}
    /**
     * Получает стоимость для конкретного типа контейнера
     */
    private function getRailCostForContainerType($containerType, $data, $isDanger = false, $ownership = 'coc'): float {
        if ($isDanger) {
            // Стоимость для опасного груза
            switch ($containerType) {
                case '20DC (<24t)':
                case '20DC':
                    return ceil(floatval($data['OPASNYY_20DC_24T'] ?? 0));
                case '20DC (24t-28t)':
                    return ceil(floatval($data['OPASNYY_20DC_24T_28T'] ?? 0));
                case '40HC (28t)':
                case '40HC':
                    return ceil(floatval($data['OPASNYY_40HC_28T'] ?? 0));
                default:
                    return 0;
            }
        } else {
            // Базовая стоимость для обычного груза
            switch ($containerType) {
                case '20DC (<24t)':
                case '20DC':
                    $key = $ownership === 'soc' ? 'DC20_24' : 'COC_20DC_24T';
                    return ceil(floatval($data[$key] ?? 0));
                case '20DC (24t-28t)':
                    $key = $ownership === 'soc' ? 'DC20_24T_28T' : 'COC_DC_24T_28T';
                    return ceil(floatval($data[$key] ?? 0));
                case '40HC (28t)':
                case '40HC':
                    $key = $ownership === 'soc' ? 'HC40_28T' : 'COC_HC_28T';
                    return ceil(floatval($data[$key] ?? 0));
                default:
                    return 0;
            }
        }
    }

    /**
     * Рассчитывает стоимость охраны для конкретного типа контейнера
     */
    private function getSecurityCostForContainerType($data, $security, $containerType): float {
        if ($security === 'no') {
            return 0;
        }
        
        // Определяем поле для охраны в зависимости от типа контейнера
        $is40HC = ($containerType === '40HC (28t)' || '40HC');
        
        //if (($security === '20' && !$is40HC) || ($security === '40' && $is40HC)) {
            $securityField = $security === '20' ? 'OKHRANA_20_FUT' : 'OKHRANA_40_FUT';
            return ceil(floatval($data[$securityField] ?? 0));
        //}
        
        return 0;
    }
/**
 * Получает стоимости для ж/д перевозок по типам контейнеров
 */
private function getRailCosts($data, $containerType, $isDanger = false): array {
    if ($isDanger) {
        // Стоимость для опасного груза
        return [
            '20' => ceil(floatval($data['OPASNYY_20DC_24'] ?? 0)),
            '20_28' => ceil(floatval($data['OPASNYY_DC20_24T_28T'] ?? 0)),
            '40' => ceil(floatval($data['OPASNYY_HC40_28T'] ?? 0))
        ];
    } else {
        // Базовая стоимость для обычного груза
        return [
            '20' => ceil(floatval($data['DC20_24'] ?? 0)),
            '20_28' => ceil(floatval($data['DC20_24T_28T'] ?? 0)),
            '40' => ceil(floatval($data['HC40_28T'] ?? 0))
        ];
    }
}

/**
 * Рассчитывает стоимость охраны
 */
private function getSecurityCost($data, $security, $containerType): float {
    if ($security === 'no') {
        return 0;
    }
    
    $securityField = ($containerType === '40HC (28t)') ? 'OKHRANA_40_FUT' : 'OKHRANA_20_FUT';
    
    if (($security === '20' && $containerType !== '40HC (28t)') || 
        ($security === '40' && $containerType === '40HC (28t)')) {
        return ceil(floatval($data[$securityField] ?? 0));
    }
    
    return 0;
}

/**
 * Создает элемент результата для ж/д перевозок
 */
private function createRailResultItem(
    $data, $containerType, $ownershipType, $hazardType, 
    $security, $profit, $baseCosts, $totalCosts, 
    $dangerBaseCosts = null, $dangerTotalCosts = null,
    $showBothOwnership = false, $showBothHazard = false,
    $alternativeOption = null
): array {
    $securityText = $security === 'no' ? 'Нет' : ($security === '20' ? '20 фут' : '40 фут');
    
    $item = [
        'rail_origin' => $data['POL'] ?? '',
        'rail_destination' => $data['POD'] ?? '',
        'rail_coc' => $containerType,
        'rail_container_ownership' => $ownershipType,
        'rail_agent' => $data['AGENT'] ?? '',
        'rail_hazard' => $hazardType,
        'rail_security' => $securityText,
        'rail_profit' => $profit,
        'cost_base_20' => $baseCosts['20'],
        'cost_base_20_28' => $baseCosts['20_28'],
        'cost_base_40' => $baseCosts['40'],
        'cost_security' => $this->getSecurityCost($data, $security, $containerType),
        'cost_total_20' => $totalCosts['20'],
        'cost_total_20_28' => $totalCosts['20_28'],
        'cost_total_40' => $totalCosts['40'],
        'show_both_ownership' => $showBothOwnership,
        'show_both_hazard' => $showBothHazard
    ];
    
    // Добавляем стоимости для опасного груза если они предоставлены
    if ($dangerBaseCosts && $dangerTotalCosts) {
        $item['cost_base_20_danger'] = $dangerBaseCosts['20'];
        $item['cost_base_20_28_danger'] = $dangerBaseCosts['20_28'];
        $item['cost_base_40_danger'] = $dangerBaseCosts['40'];
        $item['cost_total_20_danger'] = $dangerTotalCosts['20'];
        $item['cost_total_20_28_danger'] = $dangerTotalCosts['20_28'];
        $item['cost_total_40_danger'] = $dangerTotalCosts['40'];
    }
    
    // Добавляем альтернативный вариант если предоставлен
    if ($alternativeOption) {
        if ($showBothOwnership) {
            $item['soc_option'] = $alternativeOption;
        } else if ($showBothHazard) {
            $item['normal_option'] = $alternativeOption;
        }
    }
    
    return $item;
}
/**
 * Получает стоимости для морской части комбинированных перевозок по владению контейнером
 */
private function getSeaCostsForOwnership($data, $seaContainerType, $ownershipType): array {
    $is20GP = ($seaContainerType === '20DC');
    
    if ($ownershipType === 'coc') {
        $normal = $is20GP ? ($data['COC_20GP'] ?? 0) : ($data['COC_40HC'] ?? 0);
        $danger = $is20GP ? ($data['OPASNYY_20GP'] ?? 0) : ($data['OPASNYY_40HC'] ?? 0);
    } else { // soc
        $normal = $is20GP ? ($data['SOC_20GP'] ?? 0) : ($data['SOC_40HC'] ?? 0);
        $danger = $is20GP ? ($data['OPASNYY_20GP'] ?? 0) : ($data['OPASNYY_40HC'] ?? 0);
    }
    
    return [
        'normal' => !empty($normal) && floatval($normal) > 0 ? ceil(floatval($normal)) : 0,
        'danger' => !empty($danger) && floatval($danger) > 0 ? ceil(floatval($danger)) : 0
    ];
}

/**
 * Получает DROP OFF стоимость для морской части
 */
private function getSeaDropOffCost($data, $seaContainerType): float {
    if ($seaContainerType === '20DC') {
        return !empty($data['DROP_OFF_20GP']) && floatval($data['DROP_OFF_20GP']) > 0 ? ceil(floatval($data['DROP_OFF_20GP'])) : 0;
    } else {
        return !empty($data['DROP_OFF_40HC']) && floatval($data['DROP_OFF_40HC']) > 0 ? ceil(floatval($data['DROP_OFF_40HC'])) : 0;
    }
}

/**
 * Получает стоимость охраны для морской части
 */
private function getSeaSecurityCost($data, $security, $seaContainerType): float {
    if ($security === 'no') {
        return 0;
    }
    
    $securityField = ($seaContainerType === '40HC') ? 'OKHRANA_40_FUT' : 'OKHRANA_20_FUT';
    
    if (($security === '20' && $seaContainerType !== '40HC') || 
        ($security === '40' && $seaContainerType === '40HC')) {
        return !empty($data[$securityField]) && floatval($data[$securityField]) > 0 ? ceil(floatval($data[$securityField])) : 0;
    }
    
    return 0;
}

/**
 * Создает элемент результата для морской части комбинированных перевозок
 */
private function createSeaResultItemForCombined($data, $containerType, $ownership, $hazardType, 
                                                $cafPercent, $profit, $dropOffCost, 
                                                $normalCost, $dangerCost, $socNormalCost = null): array {
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
        'show_both_hazard_in_columns' => true
    ];
    
    // Расчет для обычного груза если есть стоимость
    if ($normalCost && $normalCost > 0) {
        $nettoNormal = ceil($normalCost + $dropOffCost);
        $totalNormal = ceil($nettoNormal * (1 + $cafPercent / 100) + $profit);
        
        $resultItem['cost_container_normal'] = $normalCost;
        $resultItem['cost_netto_normal'] = $nettoNormal;
        $resultItem['cost_total_normal'] = $totalNormal;
    } else {
        $resultItem['cost_container_normal'] = '-';
        $resultItem['cost_netto_normal'] = '-';
        $resultItem['cost_total_normal'] = '-';
    }
    
    // Расчет для опасного груза если есть стоимость
    if ($dangerCost && $dangerCost > 0) {
        $nettoDanger = ceil($dangerCost + $dropOffCost);
        $totalDanger = ceil($nettoDanger * (1 + $cafPercent / 100) + $profit);
        
        $resultItem['cost_container_danger'] = $dangerCost;
        $resultItem['cost_netto_danger'] = $nettoDanger;
        $resultItem['cost_total_danger'] = $dangerCost;
        $resultItem['show_both_hazard_in_columns'] = true;
    } else {
        $resultItem['cost_container_danger'] = '-';
        $resultItem['cost_netto_danger'] = '-';
        $resultItem['cost_total_danger'] = '-';
    }
    
    return $resultItem;
}

/**
 * Сопоставляет тип ж/д контейнера с морским
 */
private function mapRailToSeaContainerType($railContainerType): string {
    switch ($railContainerType) {
        case '20DC (<24t)':
        case '20DC (24t-28t)':
            return '20DC';
        case '40HC (28t)':
        case '40HC':
            return '40HC';
        default:
            return '20DC';
    }
}
/**
 * Получаем комбинированные маршруты
 *
 * @return [type]
 */
public function getCombPerevozki() {
    header('Content-Type: application/json; charset=utf-8');
    $result = [];
    $params = $_POST;
    
    try {
        $seaPol = $params['comb_sea_pol'] ?? '';
        $dropOff = $params['comb_drop_off'] ?? '';
        
        if (empty($seaPol) || empty($dropOff)) {
            echo json_encode([
                'error' => true,
                'message' => 'Не указаны обязательные параметры: порт отправления или DROP OFF'
            ], JSON_UNESCAPED_UNICODE);
            return;
        }
        
        // Получаем все морские перевозки с указанным портом отправления и DROP OFF
        $seaPerevozki = self::fetchTransportData(
            self::IBLOCK_SEA_TRANSPORTATION, 
            self::SEA_TRANSPORT_MAP,
            [
                '=NAME' => $seaPol,
                '=PROPERTY_132' => $dropOff,
            ]
        );

        if (empty($seaPerevozki)) {
            echo json_encode([
                'error' => true,
                'message' => 'Не найдены морские перевозки для порта: ' . $seaPol . ' и DROP OFF: ' . $dropOff
            ], JSON_UNESCAPED_UNICODE);
            return;
        }
        
        // Получаем порт перевалки
        $transshipmentPort = $params['comb_transshipment_port'] ?? '';
        
        // Получаем все комбинированные перевозки
        $combPerevozki = self::fetchTransportData(
            self::IBLOCK_COMBINED_TRANSPORTATION,
            self::COMB_TRANSPORT_MAP,
            [
                'PROPERTY_186' => $params['comb_rail_dest'] ?? ''
            ]
        );
        
        if (empty($combPerevozki)) {
            echo json_encode([
                'error' => true,
                'message' => 'Не найдены комбинированные перевозки'
            ], JSON_UNESCAPED_UNICODE);
            return;
        }
        
        // Получаем пункт назначения
        $combDestPoint = $params['comb_rail_dest'] ?? '';
        
        // Если выбран порт перевалки - фильтруем комбинированные перевозки
        if (!empty($transshipmentPort)) {
            $filteredCombPerevozki = [];
            foreach ($combPerevozki as $combItem) {
                // Проверяем, что PUNKT_OTPRAVLENIYA соответствует порту перевалки
                if (($combItem['PUNKT_OTPRAVLENIYA'] ?? '') === $transshipmentPort) {
                    // Если выбран пункт назначения, фильтруем и по нему
                    if (empty($combDestPoint) || ($combItem['PUNKT_NAZNACHENIYA'] ?? '') === $combDestPoint) {
                        $filteredCombPerevozki[] = $combItem;
                    }
                }
            }
            $combPerevozki = $filteredCombPerevozki;
        } elseif (!empty($combDestPoint)) {
            // Если выбран только пункт назначения без порта перевалки
            $filteredCombPerevozki = [];
            foreach ($combPerevozki as $combItem) {
                if (($combItem['PUNKT_NAZNACHENIYA'] ?? '') === $combDestPoint) {
                    $filteredCombPerevozki[] = $combItem;
                }
            }
            $combPerevozki = $filteredCombPerevozki;
        }
        
        if (empty($combPerevozki)) {
            echo json_encode([
                'error' => true,
                'message' => 'Не найдены комбинированные перевозки для указанных параметров'
            ], JSON_UNESCAPED_UNICODE);
            return;
        }

        // Параметры расчета
        $cocType = $params['comb_coc'] ?? '';
        $isHazard = ($params['comb_hazard'] ?? 'no') === 'yes';
        $security = $params['comb_security'] ?? 'no';
        $seaProfit = floatval($params['sea_profit'] ?? 0);
        $railProfit = floatval($params['rail_profit'] ?? 0);
        $containerOwnership = $params['comb_container_ownership'] ?? 'no';

        // Определяем типы контейнеров для отображения (из ж/д перевозок)
        $containerTypesToShow = empty($cocType) ? ['20DC (<24t)', '20DC (24t-28t)', '40HC (28t)'] : [$cocType];
        
        // Флаг для раздельного отображения COC/SOC
        $showBothOwnership = ($containerOwnership === 'no');

        // Для каждой найденной морской перевозки
        foreach ($seaPerevozki as $seaValue) {
            $seaPod = $seaValue['POD'] ?? '';
            $cafPercent = floatval($seaValue["CAF_KONVERT"]);

            // Фильтруем комбинированные перевозки по морскому порту прибытия
            $filteredBySeaPod = [];
            foreach ($combPerevozki as $combItem) {
                if (($combItem['PUNKT_OTPRAVLENIYA'] ?? '') === $seaPod) {
                    $filteredBySeaPod[] = $combItem;
                }
            }

            if (empty($filteredBySeaPod)) {
                continue; // Нет комбинированных перевозок для этого морского порта прибытия
            }

            // Для каждой подходящей комбинированной перевозки
            foreach ($filteredBySeaPod as $combValue) {
                $railStartStation = $combValue['STANTSIYA_OTPRAVLENIYA'] ?? '';
                $railDestStation = $combValue['STANTSIYA_NAZNACHENIYA'] ?? '';
                $combDestPoint = $combValue['PUNKT_NAZNACHENIYA'] ?? '';

                $railFilter = [
                    '=NAME' => $railStartStation,
                ];
                // Получаем ж/д перевозки для станции отправления
                $railData = self::fetchTransportData(
                    self::IBLOCK_RAIL_TRANSPORTATION, 
                    self::ZHD_TRANSPORT_MAP,
                    $railFilter
                );
                
                
                $podAgentMapping = [
                    'ВСК' => 'Восточная Стивидорная Компания',
                    'ВМРП' => 'Владивостокский морской рыбный порт',
                    'ВМПП' => 'Владивостокский морской порт "Первомайский"',
                    'ВМТП' => 'ООО "ФИТ г. Владивосток"',
                    'ППК-1' => 'ООО "ТДГ»ТK/TK действ. с 01-15/04',
                    //'Соллерс' => '',
                ];
                
    //echo json_encode([$combValue['PUNKT_OTPRAVLENIYA']], JSON_UNESCAPED_UNICODE);
    //return json_encode([$combValue], JSON_UNESCAPED_UNICODE);
                // Проверяем, есть ли соответствие для данного POD

                
                if (empty($railStartStation) || empty($railDestStation)) {
                    continue;
                }

                // Фильтруем по станции назначения
                $filteredRailData = [];
                foreach ($railData as $railItem) {

                    $railAgent = $railItem['AGENT'] ?? '';
                    if (isset($podAgentMapping[$combValue['PUNKT_OTPRAVLENIYA']])) {
                        $expectedAgent = $podAgentMapping[$combValue['PUNKT_OTPRAVLENIYA']];
                        // Если агент не соответствует ожидаемому, пропускаем эту морскую перевозку
                        if ($railAgent !== $expectedAgent) {
                            continue;
                        }
                    } else {
                        continue;
                    }

                    if (($railItem['POD'] ?? '') === $railDestStation) {
                        $filteredRailData[] = $railItem;
                    }
                }
                
                if (empty($filteredRailData)) {
                    continue; // Нет подходящих ж/д маршрутов
                }
                
                // Обрабатываем каждый тип контейнера
                foreach ($containerTypesToShow as $railContainerType) {
                    // Определяем соответствующий морской тип контейнера
                    $seaContainerType = $this->mapRailToSeaContainerType($railContainerType);
                    
                    // Для каждой подходящей ж/д перевозки
                    foreach ($filteredRailData as $railValue) {
                        // Если собственность контейнера не выбрана ('no') - показываем оба варианта в ОТДЕЛЬНЫХ РЯДАХ
                        if ($showBothOwnership) {
                            // Добавляем ряды для COC и SOC
                            $this->addCombinedRowsForBothOwnership(
                                $result, $seaValue, $railValue, $combValue,
                                $railContainerType, $seaContainerType,
                                $railStartStation, $railDestStation,
                                $cafPercent, $seaProfit, $railProfit,
                                $security, $combPerevozki, $railStartStation
                            );
                        } else {
                            // Один вариант собственности контейнера
                            $this->addCombinedRowsForSingleOwnership(
                                $result, $seaValue, $railValue, $combValue,
                                $railContainerType, $seaContainerType,
                                $railStartStation, $railDestStation,
                                $containerOwnership, $isHazard,
                                $cafPercent, $seaProfit, $railProfit,
                                $security, $combPerevozki, $railStartStation
                            );
                        }
                    }
                }
            }
        }
        
        // Если не нашли данные
        if (empty($result)) {
            $result = [
                'error' => true,
                'message' => 'Не найдены данные для указанных параметров'
            ];
        }
        
    } catch (\Exception $e) {
        file_put_contents(__DIR__ . '/error.log', date('c') . " (getCombPerevozki) " . $e->getMessage() . "\n", FILE_APPEND|LOCK_EX);
        $result = [
            'error' => true,
            'message' => 'Ошибка при получении данных: ' . $e->getMessage()
        ];
    }
    
    echo json_encode($result, JSON_UNESCAPED_UNICODE);
    return json_encode($result, JSON_UNESCAPED_UNICODE);
}

/**
 * Добавляет ряды для обоих вариантов собственности (COC и SOC)
 */
private function addCombinedRowsForBothOwnership(
    array &$result, array $seaValue, array $railValue, array $combValue,
    string $railContainerType, string $seaContainerType,
    string $railStartStation, string $railDestStation,
    float $cafPercentSea, float $seaProfit, float $railProfit,
    string $security, array $combPerevozki, string $railStartForRemark
): void {
    // Для каждого варианта собственности
    foreach (['COC', 'SOC'] as $ownership) {
        $ownershipType = strtolower($ownership);
        
        // Получаем стоимости для морской части
        $seaCosts = $this->getSeaCostsForOwnership($seaValue, $seaContainerType, $ownershipType);
        
        // Получаем DROP OFF стоимость
        $dropOffCost = $this->getSeaDropOffCost($seaValue, $seaContainerType);
        
        // Получаем стоимость охраны для морской части
        $securityCostSea = $this->getSeaSecurityCost($seaValue, $security, $seaContainerType);
        
        // Ж/Д часть - обычный груз
        $railCostNormal = $this->getRailCostForContainerType($railContainerType, $railValue, false, $ownershipType);
        $hasRailNormal = $railCostNormal > 0;
        
        // Ж/Д часть - опасный груз
        $railCostDanger = $this->getRailCostForContainerType($railContainerType, $railValue, true, $ownershipType);
        $hasRailDanger = $railCostDanger > 0;
        
        // Стоимость охраны для ЖД части
        $securityCostRail = $this->getSecurityCostForContainerType($railValue, $security, $railContainerType);
        
        // Проверяем наличие хотя бы одной стоимости
        $hasSeaCost = $seaCosts['normal'] > 0;
        $hasRailCost = $hasRailNormal || $hasRailDanger;

        if ($hasSeaCost && $hasRailCost) {
            // Итоговые стоимости ЖД
            $costRailNormal = $hasRailNormal ? ceil($railCostNormal + $securityCostRail + $railProfit) : '-';
            $costRailDanger = $hasRailDanger ? ceil($railCostDanger) : '-';
            
            // Морская часть - обычный груз
            $seaResultNormal = $this->createSeaResultItemForCombined(
                $seaValue, $seaContainerType, $ownership, 'Нет',
                $cafPercentSea, $seaProfit, $dropOffCost,
                $seaCosts['normal'] > 0 ? $seaCosts['normal'] : 0, 0,
                $ownershipType === 'soc' && $seaCosts['normal'] > 0 ? $seaCosts['normal'] : null
            );
            
            // Морская часть - опасный груз
            $seaResultDanger = $this->createSeaResultItemForCombined(
                $seaValue, $seaContainerType, $ownership, 'Да',
                $cafPercentSea, $seaProfit, $dropOffCost,
                0, $seaCosts['danger'] > 0 ? $seaCosts['danger'] : 0,
                $ownershipType === 'soc' && $seaCosts['normal'] > 0 ? $seaCosts['normal'] : null
            );
            
            // Получаем итоговые стоимости морской части
            $seaTotalNormal = isset($seaResultNormal['cost_total_normal']) && $seaResultNormal['cost_total_normal'] !== '-' 
                ? $seaResultNormal['cost_total_normal'] 
                : '-';
            
            $seaTotalDanger = isset($seaResultDanger['cost_total_danger']) && $seaResultDanger['cost_total_danger'] !== '-' 
                ? $seaResultDanger['cost_total_danger'] 
                : '-';
            
            // Формируем текстовые значения итоговых стоимостей с прочерками при отсутствии данных
            $costTotalNormalText = '';
            if ($seaTotalNormal !== '-' && $costRailNormal !== '-') {
                $costTotalNormalText = $seaTotalNormal . '$ + ' . $costRailNormal . ' руб';
            } elseif ($seaTotalNormal === '-' && $costRailNormal === '-') {
                $costTotalNormalText = '-';
            } elseif ($seaTotalNormal === '-') {
                $costTotalNormalText = '- (море) + ' . $costRailNormal . ' руб (жд)';
            } else {
                $costTotalNormalText = $seaTotalNormal . '$ (море) + - (жд)';
            }
            
            $costTotalDangerText = '';
            if ($seaTotalDanger !== '-' && $costRailDanger !== '-') {
                $costTotalDangerText = $seaTotalDanger . '$ + ' . $costRailDanger . ' руб';
            } else {
                $costTotalDangerText = '-';
            }

            // Собираем итоговый ряд
            $resultItem = [
                'comb_sea_pol' => $seaValue['POL'] ?? '',
                'comb_sea_pod' => $seaValue['POD'] ?? '',
                'comb_rail_start' => $railStartStation,
                'comb_rail_dest' => $railDestStation,
                'comb_drop_off' => $seaValue['DROP_OFF_LOCATION'] ?? '',
                'comb_transshipment_port' => $combValue['PUNKT_OTPRAVLENIYA'] ?? '',
                'comb_coc' => $railContainerType,
                'comb_container_ownership' => $ownership,
                'comb_hazard' => 'Оба варианта',
                'comb_security' => $security === 'no' ? 'Нет' : ($security === '20' ? '20 фут' : '40 фут'),
                'comb_agent' => trim(($seaValue['AGENT'] ?? '') . '; ' . ($railValue['AGENT'] ?? '')),
                'comb_remark' => $this->getCombinedRemark($seaValue, $combPerevozki, $railStartForRemark),
                
                // Морская часть - обычный груз
                'cost_sea_normal' => $seaTotalNormal !== '-' ? $seaTotalNormal : '-',
                'container_cost_normal' => $seaCosts['normal'] > 0 ? $seaCosts['normal'] : '-',
                'drop_off_cost_normal' => $dropOffCost > 0 ? $dropOffCost : '-',
                'caf_percent_normal' => $cafPercentSea,
                'security_cost_sea_normal' => $securityCostSea > 0 ? $securityCostSea : '-',
                
                // Морская часть - опасный груз
                'cost_sea_danger' => $seaTotalDanger !== '-' ? $seaTotalDanger : '-',
                'container_cost_danger' => $seaCosts['danger'] > 0 ? $seaCosts['danger'] : '-',
                'drop_off_cost_danger' => $dropOffCost > 0 ? $dropOffCost : '-',
                'caf_percent_danger' => $cafPercentSea,
                'security_cost_sea_danger' => $securityCostSea > 0 ? $securityCostSea : '-',
                
                // Ж/Д часть - обычный груз
                'cost_rail_normal' => $costRailNormal,
                'rail_base_cost_normal' => $hasRailNormal ? $railCostNormal : '-',
                'security_cost_rail_normal' => $securityCostRail > 0 ? $securityCostRail : '-',
                
                // Ж/Д часть - опасный груз
                'cost_rail_danger' => $costRailDanger,
                'rail_base_cost_danger' => $hasRailDanger ? $railCostDanger : '-',
                'security_cost_rail_danger' => $securityCostRail > 0 ? $securityCostRail : '-',
                
                // Итоговые стоимости
                'cost_total_normal_text' => $costTotalNormalText,
                'cost_total_danger_text' => $costTotalDangerText,
                
                'show_both_ownership' => false,
                'show_both_hazard_in_columns' => true
            ];
            
            $result[] = $resultItem;
        }
    }
}

/**
 * Добавляет ряды для одного варианта собственности
 */
private function addCombinedRowsForSingleOwnership(
    array &$result, array $seaValue, array $railValue, array $combValue,
    string $railContainerType, string $seaContainerType,
    string $railStartStation, string $railDestStation,
    string $containerOwnership, bool $isHazard,
    float $cafPercentSea, float $seaProfit, float $railProfit,
    string $security, array $combPerevozki, string $railStartForRemark
): void {
    $ownership = $containerOwnership === 'coc' ? 'COC' : 'SOC';
    $ownershipType = $containerOwnership;
    
    // Получаем стоимости для морской части
    $seaCosts = $this->getSeaCostsForOwnership($seaValue, $seaContainerType, $ownershipType);
    
    // Получаем DROP OFF стоимость
    $dropOffCost = $this->getSeaDropOffCost($seaValue, $seaContainerType);
    
    // Получаем стоимость охраны для морской части
    $securityCostSea = $this->getSeaSecurityCost($seaValue, $security, $seaContainerType);
    
    if (!$isHazard) {
        // Пользователь выбрал обычный груз - показываем оба варианта (обычный и опасный)
        
        // Ж/Д часть - обычный груз
        $railCostNormal = $this->getRailCostForContainerType($railContainerType, $railValue, false, $ownershipType);
        $hasRailNormal = $railCostNormal > 0;
        
        // Ж/Д часть - опасный груз
        $railCostDanger = $this->getRailCostForContainerType($railContainerType, $railValue, true, $ownershipType);
        $hasRailDanger = $railCostDanger > 0;
        
        // Стоимость охраны для ЖД части
        $securityCostRail = $this->getSecurityCostForContainerType($railValue, $security, $railContainerType);
        
        // Проверяем наличие хотя бы одной стоимости
        $hasSeaCost = $seaCosts['normal'] > 0;
        $hasRailCost = $hasRailNormal;

        if ($hasSeaCost && $hasRailCost) {
            // Итоговые стоимости ЖД
            $costRailNormal = $hasRailNormal ? ceil($railCostNormal + $securityCostRail + $railProfit) : '-';
            $costRailDanger = $hasRailDanger ? ceil($railCostDanger) : '-';
            
            // Морская часть для обычного груза
            $seaResult = $this->createSeaResultItemForCombined(
                $seaValue, $seaContainerType, $ownership, 'Нет',
                $cafPercentSea, $seaProfit, $dropOffCost,
                $seaCosts['normal'] > 0 ? $seaCosts['normal'] : 0, 
                $seaCosts['danger'] > 0 ? $seaCosts['danger'] : 0,
                $ownershipType === 'soc' && $seaCosts['normal'] > 0 ? $seaCosts['normal'] : null
            );
            
            // Получаем итоговые стоимости морской части
            $seaTotalNormal = isset($seaResult['cost_total_normal']) && $seaResult['cost_total_normal'] !== '-' 
                ? $seaResult['cost_total_normal'] 
                : '-';
            
            $seaTotalDanger = isset($seaResult['cost_total_danger']) && $seaResult['cost_total_danger'] !== '-' 
                ? $seaResult['cost_total_danger'] 
                : '-';
            
            // Формируем текстовые значения итоговых стоимостей с прочерками при отсутствии данных
            $costTotalNormalText = '';
            if ($seaTotalNormal !== '-' && $costRailNormal !== '-') {
                $costTotalNormalText = $seaTotalNormal . '$ + ' . $costRailNormal . ' руб';
            } elseif ($seaTotalNormal === '-' && $costRailNormal === '-') {
                $costTotalNormalText = '-';
            } elseif ($seaTotalNormal === '-') {
                $costTotalNormalText = '- (море) + ' . $costRailNormal . ' руб (жд)';
            } else {
                $costTotalNormalText = $seaTotalNormal . '$ (море) + - (жд)';
            }
            
            $costTotalDangerText = '';
            if ($seaTotalDanger !== '-' && $costRailDanger !== '-') {
                $costTotalDangerText = $seaTotalDanger . '$ + ' . $costRailDanger . ' руб';
            } else {
                $costTotalDangerText = '-';
            }
            
            // Собираем итоговый ряд
            $resultItem = [
                'comb_sea_pol' => $seaValue['POL'] ?? '',
                'comb_sea_pod' => $seaValue['POD'] ?? '',
                'comb_rail_start' => $railStartStation,
                'comb_rail_dest' => $railDestStation,
                'comb_drop_off' => $seaValue['DROP_OFF_LOCATION'] ?? '',
                'comb_transshipment_port' => $combValue['PUNKT_OTPRAVLENIYA'] ?? '',
                'comb_coc' => $railContainerType,
                'comb_container_ownership' => $ownership,
                'comb_hazard' => 'Нет',
                'comb_security' => $security === 'no' ? 'Нет' : ($security === '20' ? '20 фут' : '40 фут'),
                'comb_agent' => trim(($seaValue['AGENT'] ?? '') . '; ' . ($railValue['AGENT'] ?? '')),
                'comb_remark' => $this->getCombinedRemark($seaValue, $combPerevozki, $railStartForRemark),
                
                // Морская часть - обычный груз
                'cost_sea_normal' => $seaTotalNormal !== '-' ? $seaTotalNormal : '-',
                'container_cost_normal' => $seaCosts['normal'] > 0 ? $seaCosts['normal'] : '-',
                'drop_off_cost_normal' => $dropOffCost > 0 ? $dropOffCost : '-',
                'caf_percent_normal' => $cafPercentSea,
                'security_cost_sea_normal' => $securityCostSea > 0 ? $securityCostSea : '-',
                
                // Морская часть - опасный груз
                'cost_sea_danger' => $seaTotalDanger !== '-' ? $seaTotalDanger : '-',
                'container_cost_danger' => $seaCosts['danger'] > 0 ? $seaCosts['danger'] : '-',
                'drop_off_cost_danger' => $dropOffCost > 0 ? $dropOffCost : '-',
                'caf_percent_danger' => $cafPercentSea,
                'security_cost_sea_danger' => $securityCostSea > 0 ? $securityCostSea : '-',
                
                // Ж/Д часть - обычный груз
                'cost_rail_normal' => $costRailNormal,
                'rail_base_cost_normal' => $hasRailNormal ? $railCostNormal : '-',
                'security_cost_rail_normal' => $securityCostRail > 0 ? $securityCostRail : '-',
                
                // Ж/Д часть - опасный груз
                'cost_rail_danger' => $costRailDanger,
                'rail_base_cost_danger' => $hasRailDanger ? $railCostDanger : '-',
                'security_cost_rail_danger' => $securityCostRail > 0 ? $securityCostRail : '-',
                
                // Итоговые стоимости
                'cost_total_normal_text' => $costTotalNormalText,
                'cost_total_danger_text' => $costTotalDangerText,
                
                'show_both_ownership' => false,
                'show_both_hazard_in_columns' => true
            ];
            
            $result[] = $resultItem;
        }
        
    } else {
        // Пользователь выбрал опасный груз - показываем только опасный ТОЛЬКО если есть стоимость
        
        // Проверяем наличие стоимости для морской части
        if ($seaCosts['danger'] > 0) {
            // Морская часть для опасного груза
            $nettoDanger = ceil($seaCosts['danger'] + $dropOffCost);
            $totalSeaDanger = ceil($nettoDanger * (1 + $cafPercentSea / 100) + $securityCostSea + $seaProfit);
            
            // Ж/Д часть для опасного груза
            $railCostDanger = $this->getRailCostForContainerType($railContainerType, $railValue, true, $ownershipType);
            $hasRailDanger = $railCostDanger > 0;
            
            // Добавляем ряд ТОЛЬКО если есть стоимость в морской части
            // (для опасного груза морская часть обязательна)
            
            // Стоимость охраны для ЖД части
            $securityCostRail = $this->getSecurityCostForContainerType($railValue, $security, $railContainerType);
            
            // Итоговая стоимость ЖД
            $costRailDanger = $hasRailDanger ? ceil($railCostDanger) : '-';
            
            // Формируем текстовое значение итоговой стоимости с прочерками при отсутствии данных
            $costTotalText = '';
            if ($hasRailDanger) {
                $costTotalText = $totalSeaDanger . '$ + ' . $costRailDanger . ' руб';
            } else {
                $costTotalText = $totalSeaDanger . '$ (море) + - (жд)';
            }
            
            // Собираем итоговый ряд
            $resultItem = [
                'comb_sea_pol' => $seaValue['POL'] ?? '',
                'comb_sea_pod' => $seaValue['POD'] ?? '',
                'comb_rail_start' => $railStartStation,
                'comb_rail_dest' => $railDestStation,
                'comb_drop_off' => $seaValue['DROP_OFF_LOCATION'] ?? '',
                'comb_transshipment_port' => $combValue['PUNKT_OTPRAVLENIYA'] ?? '',
                'comb_coc' => $railContainerType,
                'comb_container_ownership' => $ownership,
                'comb_hazard' => 'Да',
                'comb_security' => $security === 'no' ? 'Нет' : ($security === '20' ? '20 фут' : '40 фут'),
                'comb_agent' => trim(($seaValue['AGENT'] ?? '') . '; ' . ($railValue['AGENT'] ?? '')),
                'comb_remark' => $this->getCombinedRemark($seaValue, $combPerevozki, $railStartForRemark),
                
                // Морская часть - опасный груз
                'cost_sea' => $totalSeaDanger,
                'container_cost' => $seaCosts['danger'] > 0 ? $seaCosts['danger'] : '-',
                'drop_off_cost' => $dropOffCost > 0 ? $dropOffCost : '-',
                'caf_percent' => $cafPercentSea,
                'security_cost_sea' => $securityCostSea > 0 ? $securityCostSea : '-',
                
                // Ж/Д часть - опасный груз
                'cost_rail' => $costRailDanger,
                'rail_base_cost' => $hasRailDanger ? $railCostDanger : '-',
                'security_cost_rail' => $securityCostRail > 0 ? $securityCostRail : '-',
                
                // Общая стоимость
                'cost_total_text' => $costTotalText,
                
                'show_both_ownership' => false,
                'show_both_hazard_in_columns' => false
            ];
            
            $result[] = $resultItem;
        }
    }
}
/**
 * Создает пустой элемент результата для комбинированных перевозок (прочерк)
 */
private function createEmptyCombResultItem($seaData, $containerType, $ownership, 
                                          $railStart, $railDest, $transshipmentPort, $security): array {
    return [
        'comb_sea_pol' => $seaData['POL'] ?? '',
        'comb_sea_pod' => $seaData['POD'] ?? '',
        'comb_rail_start' => $railStart,
        'comb_rail_dest' => $railDest,
        'comb_drop_off' => $seaData['DROP_OFF_LOCATION'] ?? '',
        'comb_transshipment_port' => $transshipmentPort,
        'comb_coc' => $containerType,
        'comb_container_ownership' => $ownership,
        'comb_hazard' => 'Оба варианта',
        'comb_security' => $security === 'no' ? 'Нет' : ($security === '20' ? '20 фут' : '40 фут'),
        'comb_agent' => $seaData['AGENT'] ?? '',
        'comb_remark' => $seaData['REMARK'] ?? '',
        
        // Все стоимости с прочерком
        'cost_total_normal_text' => '-',
        'cost_total_danger_text' => '-',
        
        'show_both_ownership' => false,
        'show_both_hazard_in_columns' => true,
        'empty_result' => true
    ];
}
/**
 * Находит пункт перевалки для станции отправления
 */
private function findTransshipmentPort($combPerevozki, $railStartStation): string
{
    if (empty($combPerevozki) || empty($railStartStation)) {
        return '';
    }
    
    foreach ($combPerevozki as $combItem) {
        if (trim($combItem['STANTSIYA_OTPRAVLENIYA'] ?? '') === trim($railStartStation)) {
            return $combItem['PUNKT_OTPRAVLENIYA'] ?? '';
        }
    }
    
    return '';
}

/**
 * Формирует объединенное примечание
 */
private function getCombinedRemark($seaValue, $combPerevozki, $railStartStation): string
{
    $remarks = [];
    
    // Добавляем примечание из морской перевозки
    if (!empty($seaValue['REMARK'])) {
        $remarks[] = trim($seaValue['REMARK']);
    }
    
    // Добавляем примечание из комбинированной перевозки
    if (!empty($combPerevozki)) {
        foreach ($combPerevozki as $combItem) {
            if (trim($combItem['STANTSIYA_OTPRAVLENIYA'] ?? '') === trim($railStartStation)) {
                if (!empty($combItem['REMARK'])) {
                    $remarks[] = trim($combItem['REMARK']);
                }
                break;
            }
        }
    }
    
    return implode('; ', $remarks);
}
    /**
     * Рассчитывает стоимость ж/д перевозок
     *
     * @return [type]
     */
    public function calculateRail() {
        header('Content-Type: application/json; charset=utf-8');
        $result = [];
        $params = $_POST;
        
        try {
            // Получаем данные ж/д перевозок с фильтрацией
            $zhdPerevozki = self::fetchTransportData(
                self::IBLOCK_RAIL_TRANSPORTATION, 
                self::ZHD_TRANSPORT_MAP,
                [
                    '=NAME' => $params['rail_origin'] ?? '',
                    '=PROPERTY_142' => $params['rail_destination'] ?? '',
                ]
            );
            
            // Обрабатываем каждую найденную запись
            foreach ($zhdPerevozki as $value) {
                // Определяем стоимость в зависимости от типа контейнера и опасности груза
                $cocType = $params['rail_coc'] ?? '';
                $isHazard = ($params['rail_hazard'] ?? 'no') === 'yes';
                $security = $params['rail_security'] ?? 'no';
                $profit = floatval($params['rail_profit'] ?? 0);
                
                // Базовая стоимость в зависимости от типа контейнера
                $baseCost = 0;
                
                if ($cocType === '20DC (<24t)') {
                    $baseCost = ceil($isHazard ? floatval($value['OPASNYY_20DC_24'] ?? 0) : floatval($value['DC20_24'] ?? 0));
                } elseif ($cocType === '20DC (24t-28t)') {
                    $baseCost = ceil($isHazard ? floatval($value['OPASNYY_DC20_24T_28T'] ?? 0) : floatval($value['DC20_24T_28T'] ?? 0));
                } elseif ($cocType === '40HC (28t)') {
                    $baseCost = ceil($isHazard ? floatval($value['OPASNYY_HC40_28T'] ?? 0) : floatval($value['HC40_28T'] ?? 0));
                }
                
                // Стоимость охраны
                $securityCost = 0;
                if ($security === '20') {
                    $securityCost = ceil(floatval($value['OKHRANA_20_FUT'] ?? 0));
                } elseif ($security === '40') {
                    $securityCost = ceil(floatval($value['OKHRANA_40_FUT'] ?? 0));
                }
                
                // Общая стоимость
                $totalCost = ceil($baseCost + $securityCost + $profit);
                
                // Определяем тип контейнера (COC/SOC)
                $containerOwnership = $params['rail_container_ownership'] ?? 'no';
                $containerType = 'Не выбрано';
                if ($containerOwnership === 'coc') {
                    $containerType = 'COC';
                } elseif ($containerOwnership === 'soc') {
                    $containerType = 'SOC';
                }
                
                $result[] = [
                    'rail_origin' => $value['POL'] ?? '',
                    'rail_destination' => $value['POD'] ?? '',
                    'rail_coc' => $cocType,
                    'rail_container_ownership' => $containerType,
                    'rail_agent' => $value['AGENT'] ?? '',
                    'rail_hazard' => $isHazard ? 'Да' : 'Нет',
                    'rail_security' => $security === 'no' ? 'Нет' : ($security === '20' ? '20 фут' : '40 фут'),
                    'cost_base' => $baseCost,
                    'cost_security' => $securityCost,
                    'cost_total' => $totalCost,
                    'calculation_formula' => "$baseCost (базовая) + $securityCost (охрана) + $profit (прибыль) = $totalCost"
                ];
            }
            
            // Если не нашли данные
            if (empty($result)) {
                $result = [
                    'error' => true,
                    'message' => 'Не найдены данные для указанных параметров'
                ];
            }
            
        } catch (\Exception $e) {
            file_put_contents(__DIR__ . '/error.log', date('c') . " (calculateRail) " . $e->getMessage() . "\n", FILE_APPEND|LOCK_EX);
            $result = [
                'error' => true,
                'message' => 'Ошибка при расчете: ' . $e->getMessage()
            ];
        }
        
        echo json_encode($result, JSON_UNESCAPED_UNICODE);
        return json_encode($result, JSON_UNESCAPED_UNICODE);
    }
    /**
     * Загружаем ЖД маршруты из Excel файла (старый метод)
     *
     * @return [type]
     * 
     */
    public function uploadZhd()
    {
        header('Content-Type: application/json; charset=utf-8');
        try {
            if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
                http_response_code(405);
                echo json_encode(['error' => 'Метод не разрешён'], JSON_UNESCAPED_UNICODE);
                exit;
            }
            if (empty($_FILES['file']) || $_FILES['file']['error'] !== UPLOAD_ERR_OK) {
                http_response_code(400);
                echo json_encode(['error' => 'Не удалось загрузить файл'], JSON_UNESCAPED_UNICODE);
                exit;
            }

            // читаем Excel
            try {
                $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($_FILES['file']['tmp_name']);
                $sheet = $spreadsheet->getActiveSheet();
                // разлепляем объединенные ячейки, копируя значение
                foreach ($sheet->getMergeCells() as $range) {
                    $cells = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::extractAllCellReferencesInRange($range);
                    if (!$cells) continue;
                    $first = $cells[0];
                    $value = $sheet->getCell($first)->getValue();
                    foreach ($cells as $coord) {
                        $sheet->getCell($coord)->setValue($value);
                    }
                }
                $rows = $sheet->toArray(null, true, true, true);
            } catch (\Throwable $e) {
                file_put_contents(__DIR__ . '/error.log', date('c') . " (Excel ZHD) " . $e->getMessage() . "\n", FILE_APPEND | LOCK_EX);
                http_response_code(500);
                echo json_encode(['error' => 'Не удалось прочитать Excel: ' . $e->getMessage()], JSON_UNESCAPED_UNICODE);
                exit;
            }
            if (empty($rows) || empty($rows[1])) {
                http_response_code(422);
                echo json_encode(['error' => 'В файле нет данных или заголовка'], JSON_UNESCAPED_UNICODE);
                exit;
            }
            $headerRow = $rows[1];
            $cols = ['A','B','C','D','E','F','G','H','I','J','K', 'L', 'M', 'N']; // до K включительно
            $added = 0;
            $errors = [];
            foreach ($rows as $idx => $row) {
                // пропускаем заголовок и его возможные повторения
                $isHeader = ($idx === 1);
                if (!$isHeader) {
                    $same = true;
                    foreach ($cols as $c) {
                        if (trim((string)($row[$c] ?? '')) !== trim((string)($headerRow[$c] ?? ''))) {
                            $same = false; break;
                        }
                    }
                    $isHeader = $same;
                }
                if ($isHeader) continue;

                // пропускаем полностью пустые строки
                $allEmpty = true;
                foreach ($cols as $c) {
                    if (trim((string)($row[$c] ?? '')) !== '') { $allEmpty = false; break; }
                }
                if ($allEmpty) continue;

                try {
                    $response = \CRest::call('lists.element.add', [
                        'IBLOCK_TYPE_ID' => 'lists',
                        'IBLOCK_ID'      => 30,
                        'ELEMENT_CODE'   => 'el_' . $idx . rand(1, 9999),
                        'FIELDS'         => [
                            'NAME'         => trim((string)$row['A']),
                            'PROPERTY_142' => trim((string)$row['B']),
                            'PROPERTY_212' => str_replace(',', '', trim((string)$row['C'])),
                            'PROPERTY_214' => str_replace(',', '', trim((string)$row['D'])),
                            'PROPERTY_216' => str_replace(',', '', trim((string)$row['E'])),
                            'PROPERTY_166' => str_replace(',', '', trim((string)$row['F'])),
                            'PROPERTY_170' => str_replace(',', '', trim((string)$row['G'])),
                            'PROPERTY_174' => str_replace(',', '', trim((string)$row['H'])),
                            'PROPERTY_168' => str_replace(',', '', trim((string)$row['I'])),
                            'PROPERTY_172' => str_replace(',', '', trim((string)$row['J'])),
                            'PROPERTY_176' => str_replace(',', '', trim((string)$row['K'])),
                            'PROPERTY_178' => str_replace(',', '', trim((string)$row['L'])),
                            'PROPERTY_180' => str_replace(',', '', trim((string)$row['M'])),
                            'PROPERTY_196' => trim((string)$row['N']),
                        ],
                    ]);

                    if (!isset($response['result'])) {
                        $errors[] = ['row' => $idx, 'error' => $response['error_description'] ?? 'Неизвестная ошибка Bitrix24'];
                    } else {
                        $added++;
                    }
                } catch (\Throwable $e) {
                    $errors[] = ['row' => $idx, 'error' => $e->getMessage()];
                }
            }
            http_response_code($errors ? 207 : 200);
            echo json_encode(
                ['result' => $errors === [], 'added' => $added, 'errors' => $errors, 'message' => 'Загрузка ЖД завершена'],
                JSON_UNESCAPED_UNICODE
            );
            exit;
        } catch (\Throwable $e) {
            file_put_contents(__DIR__ . '/error.log', date('c') . " (uploadZhd) " . $e->getMessage() . "\n", FILE_APPEND | LOCK_EX);
            http_response_code(500);
            echo json_encode(['error' => 'Серверная ошибка: ' . $e->getMessage()], JSON_UNESCAPED_UNICODE);
            exit;
        }
    }

    /**
     * Загружаем морские маршруты из Excel файла (старый метод)
     *
     * @return [type]
     * 
     */
    public function uploadSea()
    {
        header('Content-Type: application/json; charset=utf-8');
        try {
            if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
                http_response_code(405);
                echo json_encode(['error' => 'Метод не разрешён'], JSON_UNESCAPED_UNICODE);
                exit;
            }
            if (empty($_FILES['file']) || $_FILES['file']['error'] !== UPLOAD_ERR_OK) {
                http_response_code(400);
                echo json_encode(['error' => 'Не удалось загрузить файл'], JSON_UNESCAPED_UNICODE);
                exit;
            }

            // читаем Excel
            try {
                $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($_FILES['file']['tmp_name']);
                $sheet = $spreadsheet->getActiveSheet();

                // разлепляем объединенные ячейки, копируя значение
                foreach ($sheet->getMergeCells() as $range) {
                    $cells = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::extractAllCellReferencesInRange($range);
                    if (!$cells) continue;
                    $first = $cells[0];
                    $value = $sheet->getCell($first)->getValue();
                    foreach ($cells as $coord) {
                        $sheet->getCell($coord)->setValue($value);
                    }
                }

                $rows = $sheet->toArray(null, true, true, true);
                echo json_encode(
                ['result' => $rows],
                JSON_UNESCAPED_UNICODE
            );
            } catch (\Throwable $e) {
                file_put_contents(__DIR__ . '/error.log', date('c') . " (Excel SEA) " . $e->getMessage() . "\n", FILE_APPEND | LOCK_EX);
                http_response_code(500);
                echo json_encode(['error' => 'Не удалось прочитать Excel: ' . $e->getMessage()], JSON_UNESCAPED_UNICODE);
                exit;
            }

            if (empty($rows) || empty($rows[1])) {
                http_response_code(422);
                echo json_encode(['error' => 'В файле нет данных или заголовка'], JSON_UNESCAPED_UNICODE);
                exit;
            }

            $headerRow = $rows[1];
            $cols = ['A','B','C','D','E','F','G','H','I','J', 'K', 'L', 'M', 'N']; // до J включительно
            $added = 0;
            $errors = [];

            foreach ($rows as $idx => $row) {
                // пропускаем заголовок и его возможные повторения
                $isHeader = ($idx === 1);
                if (!$isHeader) {
                    $same = true;
                    foreach ($cols as $c) {
                        if (trim((string)($row[$c] ?? '')) !== trim((string)($headerRow[$c] ?? ''))) {
                            $same = false; break;
                        }
                    }
                    $isHeader = $same;
                }
                if ($isHeader) continue;

                // пропускаем полностью пустые строки
                $allEmpty = true;
                foreach ($cols as $c) {
                    if (trim((string)($row[$c] ?? '')) !== '') { $allEmpty = false; break; }
                }
                if ($allEmpty) continue;

                try {
                    $response = \CRest::call('lists.element.add', [
                        'IBLOCK_TYPE_ID' => 'lists',
                        'IBLOCK_ID'      => 28,
                        'ELEMENT_CODE'   => 'el_' . $idx . rand(1, 9999),
                        'FIELDS' => [
                            'NAME'         => trim((string)$row['A']),  // Порт
                            'PROPERTY_126' => trim((string)$row['B']),
                            'PROPERTY_162' => str_replace(',', '', trim((string)$row['C'])),
                            'PROPERTY_164' => str_replace(',', '', trim((string)$row['D'])),
                            'PROPERTY_132' => str_replace(',', '', trim((string)$row['E'])),
                            'PROPERTY_134' => str_replace(',', '', trim((string)$row['F'])),
                            'PROPERTY_136' => str_replace(',', '', trim((string)$row['G'])),
                            'PROPERTY_138' => str_replace(',', '', trim((string)$row['H'])),
                            'PROPERTY_140' => trim((string)$row['I']),
                            'PROPERTY_192' => str_replace(',', '', trim((string)$row['J'])),
                            'PROPERTY_202' => str_replace(',', '', trim((string)$row['K'])),
                            'PROPERTY_200' => str_replace(',', '', trim((string)$row['L'])),
                            'PROPERTY_208' => trim((string)$row['M']),
                            'PROPERTY_210' => trim((string)$row['N']), // agent

                        ],
                    ]);

                    if (!isset($response['result'])) {
                        $errors[] = ['row' => $idx, 'error' => $response['error_description'] ?? 'Неизвестная ошибка Bitrix24'];
                    } else {
                        $added++;
                    }
                } catch (\Throwable $e) {
                    $errors[] = ['row' => $idx, 'error' => $e->getMessage()];
                }
            }

            http_response_code($errors ? 207 : 200);
            echo json_encode(
                ['result' => $errors === [], 'added' => $added, 'errors' => $errors, 'message' => 'Загрузка морских маршрутов завершена'],
                JSON_UNESCAPED_UNICODE
            );
            exit;

        } catch (\Throwable $e) {
            file_put_contents(__DIR__ . '/error.log', date('c') . " (uploadSea) " . $e->getMessage() . "\n", FILE_APPEND | LOCK_EX);
            http_response_code(500);
            echo json_encode(['error' => 'Серверная ошибка: ' . $e->getMessage()], JSON_UNESCAPED_UNICODE);
            exit;
        }
    }

    /**
     * Загружаем комбинированные маршруты из Excel файла (старый метод)
     *
     * @return [type]
     * 
     */
    public function uploadComb()
    {
        header('Content-Type: application/json; charset=utf-8');
        // Общий перехват любых фатальных ошибок/исключений
        try {
            if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
                http_response_code(405);
                echo json_encode(['error' => 'Метод не разрешён'], JSON_UNESCAPED_UNICODE);
                exit;
            }
            if (empty($_FILES['file']) || $_FILES['file']['error'] !== UPLOAD_ERR_OK) {
                http_response_code(400);
                echo json_encode(['error' => 'Не удалось загрузить файл'], JSON_UNESCAPED_UNICODE);
                exit;
            }
            // (опционально) проверка размера / типа
            if ($_FILES['file']['size'] <= 0) {
                http_response_code(400);
                echo json_encode(['error' => 'Пустой файл'], JSON_UNESCAPED_UNICODE);
                exit;
            }

            // Чтение Excel
            try {
                $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($_FILES['file']['tmp_name']);
                $sheet = $spreadsheet->getActiveSheet();
                // Разлепить объединённые ячейки, проставив значения
                foreach ($sheet->getMergeCells() as $range) {
                    $cells = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::extractAllCellReferencesInRange($range);
                    if (!$cells) continue;
                    $first = $cells[0];
                    $value = $sheet->getCell($first)->getValue();
                    foreach ($cells as $coord) {
                        $sheet->getCell($coord)->setValue($value);
                    }
                }
                // Получаем массив с буквенными ключами столбцов
                $rows = $sheet->toArray(null, true, true, true);
            } catch (\Throwable $e) {
                file_put_contents(__DIR__ . '/error.log', date('c') . " (Excel) " . $e->getMessage() . "\n", FILE_APPEND | LOCK_EX);
                http_response_code(500);
                echo json_encode(['error' => 'Не удалось прочитать Excel: ' . $e->getMessage()], JSON_UNESCAPED_UNICODE);
                exit;
            }

            if (empty($rows) || empty($rows[1])) {
                http_response_code(422);
                echo json_encode(['error' => 'В файле нет данных или заголовка'], JSON_UNESCAPED_UNICODE);
                exit;
            }

            $headerRow = $rows[1];
            $added = 0;
            $errors = [];
            foreach ($rows as $idx => $row) {
                // пропускаем заголовок и его возможные повторения
                if (
                    $idx === 1 ||
                    (
                        trim((string)$row['A']) === trim((string)$headerRow['A']) &&
                        trim((string)$row['B']) === trim((string)$headerRow['B']) &&
                        trim((string)$row['C']) === trim((string)$headerRow['C']) &&
                        trim((string)$row['D']) === trim((string)$headerRow['D']) &&
                        trim((string)$row['E']) === trim((string)$headerRow['E']) &&
                        trim((string)$row['F']) === trim((string)$headerRow['F'])
                    )
                ) {
                    continue;
                }
                try {
                    $response = \CRest::call('lists.element.add', [
                        'IBLOCK_TYPE_ID' => 'lists',
                        'IBLOCK_ID'      => 32,
                        'ELEMENT_CODE'   => 'el_' . $idx . rand(1, 9999),
                        'FIELDS'         => [
                            'NAME'         => trim((string)$row['A']), // Порт
                            'PROPERTY_182' => trim((string)$row['B']), // Пункт отправления
                            'PROPERTY_184' => trim((string)$row['C']), // Станция отправления
                            'PROPERTY_186' => trim((string)$row['D']), // Пункт назначения
                            'PROPERTY_188' => trim((string)$row['E']), // Станция назначения
                            'PROPERTY_190' => trim((string)$row['F']), // Remark
                        ],
                    ]);

                    if (!isset($response['result'])) {
                        $errors[] = ['row' => $idx, 'error' => $response['error_description'] ?? 'Неизвестная ошибка Bitrix24'];
                    } else {
                        $added++;
                    }
                } catch (\Throwable $e) {
                    $errors[] = ['row' => $idx, 'error' => $e->getMessage()];
                }
            }

            // Если были ошибки по строкам — вернём 207 Multi-Status, иначе 200
            if ($errors) {
                http_response_code(207);
            } else {
                http_response_code(200);
            }

            echo json_encode(
                ['result' => $errors === [], 'added' => $added, 'errors' => $errors, 'message' => 'Загрузка завершена'],
                JSON_UNESCAPED_UNICODE
            );
            exit;
        } catch (\Throwable $e) {
            // Фатал вне наших try/catch — лог и 500
            file_put_contents(__DIR__ . '/error.log', date('c') . " (uploadComb) " . $e->getMessage() . "\n", FILE_APPEND | LOCK_EX);
            http_response_code(500);
            echo json_encode(['error' => 'Серверная ошибка: ' . $e->getMessage()], JSON_UNESCAPED_UNICODE);
            exit;
        }
    }

    /**
     * Загружаем ЖД маршруты из JSON данных (новый метод для Vue.js)
     * Принимает JSON данные, отправленные с клиента
     *
     * @return void
     */
    public function uploadZhdData()
    {
        header('Content-Type: application/json; charset=utf-8');
        try {
            if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
                http_response_code(405);
                echo json_encode(['error' => 'Метод не разрешён'], JSON_UNESCAPED_UNICODE);
                return;
            }

            $input = json_decode(file_get_contents('php://input'), true);
            
            if (empty($input) || !is_array($input)) {
                http_response_code(400);
                echo json_encode(['error' => 'Нет данных для загрузки'], JSON_UNESCAPED_UNICODE);
                return;
            }

            $added = 0;
            $errors = [];

            foreach ($input as $idx => $row) {
                try {
                    $response = \CRest::call('lists.element.add', [
                        'IBLOCK_TYPE_ID' => 'lists',
                        'IBLOCK_ID'      => 30,
                        'ELEMENT_CODE'   => 'el_' . $idx . rand(1, 9999),
                        'FIELDS'         => [
                            'NAME'         => trim((string)($row['A'] ?? '')),
                            'PROPERTY_142' => trim((string)($row['B'] ?? '')),
                            'PROPERTY_212' => str_replace(',', '', trim((string)($row['C'] ?? ''))),
                            'PROPERTY_214' => str_replace(',', '', trim((string)($row['D'] ?? ''))),
                            'PROPERTY_216' => str_replace(',', '', trim((string)($row['E'] ?? ''))),
                            'PROPERTY_166' => str_replace(',', '', trim((string)($row['F'] ?? ''))),
                            'PROPERTY_170' => str_replace(',', '', trim((string)($row['G'] ?? ''))),
                            'PROPERTY_174' => str_replace(',', '', trim((string)($row['H'] ?? ''))),
                            'PROPERTY_168' => str_replace(',', '', trim((string)($row['I'] ?? ''))),
                            'PROPERTY_172' => str_replace(',', '', trim((string)($row['J'] ?? ''))),
                            'PROPERTY_176' => str_replace(',', '', trim((string)($row['K'] ?? ''))),
                            'PROPERTY_178' => str_replace(',', '', trim((string)($row['L'] ?? ''))),
                            'PROPERTY_180' => str_replace(',', '', trim((string)($row['M'] ?? ''))),
                            'PROPERTY_196' => trim((string)($row['N'] ?? '')),
                        ],
                    ]);

                    if (!isset($response['result'])) {
                        $errors[] = ['row' => $idx, 'error' => $response['error_description'] ?? 'Неизвестная ошибка Bitrix24'];
                    } else {
                        $added++;
                    }
                } catch (\Throwable $e) {
                    $errors[] = ['row' => $idx, 'error' => $e->getMessage()];
                }
            }

            http_response_code($errors ? 207 : 200);
            echo json_encode(
                ['result' => $errors === [], 'added' => $added, 'errors' => $errors, 'message' => 'Загрузка ЖД завершена'],
                JSON_UNESCAPED_UNICODE
            );
            return;
        } catch (\Throwable $e) {
            file_put_contents(__DIR__ . '/error.log', date('c') . " (uploadZhdData) " . $e->getMessage() . "\n", FILE_APPEND | LOCK_EX);
            http_response_code(500);
            echo json_encode(['error' => 'Серверная ошибка: ' . $e->getMessage()], JSON_UNESCAPED_UNICODE);
            return;
        }
    }

    /**
     * Загружаем морские маршруты из JSON данных (новый метод для Vue.js)
     * Принимает JSON данные, отправленные с клиента
     *
     * @return void
     */
    public function uploadSeaData()
    {
        header('Content-Type: application/json; charset=utf-8');
        try {
            if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
                http_response_code(405);
                echo json_encode(['error' => 'Метод не разрешён'], JSON_UNESCAPED_UNICODE);
                return;
            }

            $input = json_decode(file_get_contents('php://input'), true);
            
            if (empty($input) || !is_array($input)) {
                http_response_code(400);
                echo json_encode(['error' => 'Нет данных для загрузки'], JSON_UNESCAPED_UNICODE);
                return;
            }

            $added = 0;
            $errors = [];

            foreach ($input as $idx => $row) {
                try {
                    $response = \CRest::call('lists.element.add', [
                        'IBLOCK_TYPE_ID' => 'lists',
                        'IBLOCK_ID'      => 28,
                        'ELEMENT_CODE'   => 'el_' . $idx . rand(1, 9999),
                        'FIELDS' => [
                            'NAME'         => trim((string)($row['A'] ?? '')),  // Порт
                            'PROPERTY_126' => trim((string)($row['B'] ?? '')),
                            'PROPERTY_162' => str_replace(',', '', trim((string)($row['C'] ?? ''))),
                            'PROPERTY_164' => str_replace(',', '', trim((string)($row['D'] ?? ''))),
                            'PROPERTY_132' => str_replace(',', '', trim((string)($row['E'] ?? ''))),
                            'PROPERTY_134' => str_replace(',', '', trim((string)($row['F'] ?? ''))),
                            'PROPERTY_136' => str_replace(',', '', trim((string)($row['G'] ?? ''))),
                            'PROPERTY_138' => str_replace(',', '', trim((string)($row['H'] ?? ''))),
                            'PROPERTY_140' => trim((string)($row['I'] ?? '')),
                            'PROPERTY_192' => str_replace(',', '', trim((string)($row['J'] ?? ''))),
                            'PROPERTY_202' => str_replace(',', '', trim((string)($row['K'] ?? ''))),
                            'PROPERTY_200' => str_replace(',', '', trim((string)($row['L'] ?? ''))),
                            'PROPERTY_208' => trim((string)($row['M'] ?? '')),
                            'PROPERTY_210' => trim((string)($row['N'] ?? '')), // agent
                        ],
                    ]);

                    if (!isset($response['result'])) {
                        $errors[] = ['row' => $idx, 'error' => $response['error_description'] ?? 'Неизвестная ошибка Bitrix24'];
                    } else {
                        $added++;
                    }
                } catch (\Throwable $e) {
                    $errors[] = ['row' => $idx, 'error' => $e->getMessage()];
                }
            }

            http_response_code($errors ? 207 : 200);
            echo json_encode(
                ['result' => $errors === [], 'added' => $added, 'errors' => $errors, 'message' => 'Загрузка морских маршрутов завершена'],
                JSON_UNESCAPED_UNICODE
            );
            return;
        } catch (\Throwable $e) {
            file_put_contents(__DIR__ . '/error.log', date('c') . " (uploadSeaData) " . $e->getMessage() . "\n", FILE_APPEND | LOCK_EX);
            http_response_code(500);
            echo json_encode(['error' => 'Серверная ошибка: ' . $e->getMessage()], JSON_UNESCAPED_UNICODE);
            return;
        }
    }

    /**
     * Загружаем комбинированные маршруты из JSON данных (новый метод для Vue.js)
     * Принимает JSON данные, отправленные с клиента
     *
     * @return void
     */
    public function uploadCombData()
    {
        header('Content-Type: application/json; charset=utf-8');
        try {
            if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
                http_response_code(405);
                echo json_encode(['error' => 'Метод не разрешён'], JSON_UNESCAPED_UNICODE);
                return;
            }

            $input = json_decode(file_get_contents('php://input'), true);
            
            if (empty($input) || !is_array($input)) {
                http_response_code(400);
                echo json_encode(['error' => 'Нет данных для загрузки'], JSON_UNESCAPED_UNICODE);
                return;
            }

            $added = 0;
            $errors = [];

            foreach ($input as $idx => $row) {
                try {
                    $response = \CRest::call('lists.element.add', [
                        'IBLOCK_TYPE_ID' => 'lists',
                        'IBLOCK_ID'      => 32,
                        'ELEMENT_CODE'   => 'el_' . $idx . rand(1, 9999),
                        'FIELDS'         => [
                            'NAME'         => trim((string)($row['A'] ?? '')), // Порт
                            'PROPERTY_182' => trim((string)($row['B'] ?? '')), // Пункт отправления
                            'PROPERTY_184' => trim((string)($row['C'] ?? '')), // Станция отправления
                            'PROPERTY_186' => trim((string)($row['D'] ?? '')), // Пункт назначения
                            'PROPERTY_188' => trim((string)($row['E'] ?? '')), // Станция назначения
                            'PROPERTY_190' => trim((string)($row['F'] ?? '')), // Remark
                        ],
                    ]);

                    if (!isset($response['result'])) {
                        $errors[] = ['row' => $idx, 'error' => $response['error_description'] ?? 'Неизвестная ошибка Bitrix24'];
                    } else {
                        $added++;
                    }
                } catch (\Throwable $e) {
                    $errors[] = ['row' => $idx, 'error' => $e->getMessage()];
                }
            }

            http_response_code($errors ? 207 : 200);
            echo json_encode(
                ['result' => $errors === [], 'added' => $added, 'errors' => $errors, 'message' => 'Загрузка комбинированных маршрутов завершена'],
                JSON_UNESCAPED_UNICODE
            );
            return;
        } catch (\Throwable $e) {
            file_put_contents(__DIR__ . '/error.log', date('c') . " (uploadCombData) " . $e->getMessage() . "\n", FILE_APPEND | LOCK_EX);
            http_response_code(500);
            echo json_encode(['error' => 'Серверная ошибка: ' . $e->getMessage()], JSON_UNESCAPED_UNICODE);
            return;
        }
    }

    /**
     * Универсальная функция для получения и маппинга данных из списка.
     *
     * @param int   $iblockId  ID инфоблока (списка)
     * @param array $map       Ассоциативный массив [старый_ключ => новый_ключ]
     * 
     * @return array ИТОГ
     */
    private static function fetchTransportData(int $iblockId, array $map, $filter = []): array
{
    $allElements = [];
    $pageSize = 50;
    $page = 0;
    
    do {
        $response = CRest::call('lists.element.get', [
            'IBLOCK_TYPE_ID' => 'lists',
            'IBLOCK_ID' => $iblockId,
            'FILTER' => $filter,
            'start' => $pageSize * $page,
        ]);

        if (isset($response['result']) && is_array($response['result'])) {
            $elements = $response['result'];
            
            // Добавляем элементы в общий массив
            $allElements = array_merge($allElements, $elements);
            
            // Проверяем, нужно ли делать следующий запрос
            if (count($elements) < $pageSize) {
                break; // Получены все элементы с этой страницы
            }
            
            $page++;
            
        } else {
            // Логируем ошибку
            file_put_contents(__DIR__ . '/fetch_transport_data_error.log', 
                "Error on page {$page}: " . json_encode($response, JSON_UNESCAPED_UNICODE) . PHP_EOL, 
                FILE_APPEND | LOCK_EX);
            break;
        }
        
    } while (true);
    
    // Преобразуем все полученные элементы согласно карте
    $result = array_map(function(array $item) use ($map) {
        $row = [];
        foreach ($map as $oldKey => $newKey) {
            if (!array_key_exists($oldKey, $item)) {
                continue;
            }
            $value = $item[$oldKey];
            // если значение — массив, берем первый элемент
            if (is_array($value)) {
                $value = reset($value);
            }
            $row[$newKey] = $value;
        }
        return $row;
    }, $allElements);

    return $result;
}

    /**
     * Экшен install: регистрирует локальное приложение в Битрикс24,
     * используя REST-метод placement.bind.
     */
    public function install()
    {
        // Собираем все входящие данные, включая параметры авторизации Bitrix24
        $params = $_REQUEST;
        $domain = $params['DOMAIN'] ?? '';
        $newAccessToken  = $params['AUTH_ID'] ?? '';
        $newRefreshToken = $params['REFRESH_ID'] ?? '';
        file_put_contents(__DIR__ . '/data_stat.log', date('c') . ": (Пришедшие данные) " . json_encode($params) . PHP_EOL, FILE_APPEND|LOCK_EX);

        // Загружаем конфигурацию приложения
        $configFile = __DIR__ . '/app_config.php';
        if (!file_exists($configFile)) {
            return;
        }
        $config = include $configFile;
        
        if (!$domain || !$newAccessToken) {
            header('HTTP/1.0 400 Bad Request');
            echo 'Отсутствуют параметры авторизации.';
            return;
        }

        // Сохраняем новые токены в конфиг
        $config['access_token']  = $newAccessToken;
        $config['refresh_token'] = $newRefreshToken;
        $export = var_export($config, true);
        $phpCode = "<?php\nreturn {$export};\n";
        if (false === file_put_contents($configFile, $phpCode, LOCK_EX)) {
            header('HTTP/1.0 500 Internal Server Error');
            return;
        }

        // запишет settings.json
        $result = CRest::installApp();  
        if (!empty($result['error'])) {
            echo "ошибка регистрации";
            return;
        }
        file_put_contents(__DIR__ . '/data_stat.log', date('c') . ": (итог регистрации) " . json_encode($result) . PHP_EOL, FILE_APPEND|LOCK_EX);


        // Передаём $result в файл install.php
        // Путь относительно текущего контроллера:
        $viewFile = __DIR__ . '/crestV136/install.php';
        if (file_exists($viewFile)) {
            require $viewFile;
        } else {
            echo "Не найден файл представления install.php";
        }

        // //Регистрируем обработчик встраивания
        // $result = CRest::call('placement.bind', [
        //     'auth' => $newAccessToken,
        //     'PLACEMENT'   => 'LEFT_MENU',
        //     'HANDLER'     => $config['handler'],
        //     'TITLE'       => $config['title'],
        //     'DESCRIPTION' => $config['description'],
        //     //'LANG_ALL'  => ['ru' => $config['title']],
        //     // 'ADDITIONAL'=> []   // дополнительные параметры
        // ]);
    }
}

// Точка входа
$action = isset($_GET['action']) ? $_GET['action'] : 'index';
$controller = new TransportationCalculatorController();
if (method_exists($controller, $action)) {
    $controller->{$action}();
} else {
    header('HTTP/1.0 404 Not Found');
    echo 'Экшен ' . htmlspecialchars($action) . ' не найден.';
}
