<?php

// Подключаем PHP-SDK CRest для работы с REST API
require_once __DIR__ . '/crestV136/crest.php';

// Подключаем Composer
require_once __DIR__ . '/vendor/autoload.php';

// Подключаем библиотеку разбора Excel
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

/**
 * Контроллер для обработки запросов калькулятора стоимости перевозок.
 * Обрабатывает два экшена: index (отображает форму) и install (устанавливает локальное приложение).
 */
class TransportationCalculatorController
{
    // маппинг полей жд перевозок
    public const ZHD_TRANSPORT_MAP = [
        'NAME' => 'POL',
        'PROPERTY_142' => 'POD',
        'PROPERTY_166' => 'DC20_24',
        'PROPERTY_168' => 'OPASNYY_20DC_24',
        'PROPERTY_170' => 'DC20_24T_28T',
        'PROPERTY_172' => 'OPASNYY_DC20_24T_28T',
        'PROPERTY_174' => 'HC40_28T',
        'PROPERTY_176' => 'OPASNYY_HC40_28T',
        'PROPERTY_178' => 'OKHRANA_20_FUT',
        'PROPERTY_180' => 'OKHRANA_40_FUT',
        'PROPERTY_196' => 'AGENT'
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
        'PROPERTY_208' => 'OPASNYY_20DC',
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
    
    /**
     * Экшен index: отрисовывает форму расчета.
     */
    public function index()
    {
        // данные жд перевозок
        $zhdPerevozki  = self::fetchTransportData(30, self::ZHD_TRANSPORT_MAP);
        // данные морских перевозок
        $seaPerevozki  = self::fetchTransportData(28, self::SEA_TRANSPORT_MAP);
        // данные комбинированных перевозок
        $combPerevozki = self::fetchTransportData(32, self::COMB_TRANSPORT_MAP);

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
     * Экспорт результатов морских перевозок в Excel
     */
    public function exportSeaToExcel()
    {
        $data = json_decode(file_get_contents('php://input'), true);
        
        if (empty($data) || !is_array($data)) {
            header('Content-Type: application/json; charset=utf-8');
            echo json_encode(['error' => true, 'message' => 'Нет данных для экспорта']);
            return;
        }
        
        $this->generateExcel($data, 'sea_export_' . date('Y-m-d_H-i-s'));
    }
    
    /**
     * Экспорт результатов ж/д перевозок в Excel
     */
    public function exportRailToExcel()
    {
        $data = json_decode(file_get_contents('php://input'), true);
        
        if (empty($data) || !is_array($data)) {
            header('Content-Type: application/json; charset=utf-8');
            echo json_encode(['error' => true, 'message' => 'Нет данных для экспорта']);
            return;
        }
        
        $this->generateExcel($data, 'rail_export_' . date('Y-m-d_H-i-s'), 'rail');
    }
    
    /**
     * Экспорт результатов комбинированных перевозок в Excel
     */
    public function exportCombToExcel()
    {
        $data = json_decode(file_get_contents('php://input'), true);
        
        if (empty($data) || !is_array($data)) {
            header('Content-Type: application/json; charset=utf-8');
            echo json_encode(['error' => true, 'message' => 'Нет данных для экспорта']);
            return;
        }
        
        $this->generateExcel($data, 'combined_export_' . date('Y-m-d_H-i-s'), 'combined');
    }
    
    /**
     * Генерация Excel файла
     */
    private function generateExcel($data, $filename, $type = 'sea')
    {
        try {
            $spreadsheet = new Spreadsheet();
            $sheet = $spreadsheet->getActiveSheet();
            
            // Заголовки в зависимости от типа данных
            if ($type === 'sea') {
                $headers = [
                    'A' => 'Порт отправления (POL)',
                    'B' => 'Порт прибытия (POD)',
                    'C' => 'DROP OFF LOCATION',
                    'D' => 'Тип контейнера',
                    'E' => 'Собственность контейнера',
                    'F' => 'Опасный груз',
                    'G' => 'Охрана',
                    'H' => 'Стоимость контейнера ($)',
                    'I' => 'Стоимость DROP OFF ($)',
                    'J' => 'Стоимость охраны ($)',
                    'K' => 'NETTO ($)',
                    'L' => 'CAF (%)',
                    'M' => 'Profit ($)',
                    'N' => 'Итоговая стоимость ($)',
                    'O' => 'Итоговая стоимость (опасный) ($)',
                    'P' => 'Агент',
                    'Q' => 'Примечание',
                ];
                
                // Заполняем заголовки
                foreach ($headers as $col => $header) {
                    $sheet->setCellValue($col . '1', $header);
                    $sheet->getStyle($col . '1')->getFont()->setBold(true);
                }
                
                // Заполняем данные
                $row = 2;
                foreach ($data as $item) {
                    $sheet->setCellValue('A' . $row, $item['sea_pol'] ?? '');
                    $sheet->setCellValue('B' . $row, $item['sea_pod'] ?? '');
                    $sheet->setCellValue('C' . $row, $item['sea_drop_off_location'] ?? '');
                    $sheet->setCellValue('D' . $row, $item['sea_coc'] ?? '');
                    $sheet->setCellValue('E' . $row, $item['sea_container_ownership'] ?? '');
                    $sheet->setCellValue('F' . $row, $item['sea_hazard'] ?? 'Нет');
                    $sheet->setCellValue('G' . $row, $item['sea_security'] ?? 'Нет');
                    $sheet->setCellValue('H' . $row, $item['cost_container'] ?? 0);
                    $sheet->setCellValue('I' . $row, $item['cost_drop_off'] ?? 0);
                    $sheet->setCellValue('J' . $row, $item['cost_security'] ?? 0);
                    $sheet->setCellValue('K' . $row, $item['cost_netto'] ?? 0);
                    $sheet->setCellValue('L' . $row, $item['sea_caf_percent'] ?? 0);
                    $sheet->setCellValue('M' . $row, $item['sea_profit'] ?? 0);
                    $sheet->setCellValue('N' . $row, $item['cost_total'] ?? 0);
                    $sheet->setCellValue('O' . $row, $item['cost_total_danger'] ?? ($item['cost_total'] ?? 0));
                    $sheet->setCellValue('P' . $row, $item['sea_agent'] ?? '');
                    $sheet->setCellValue('Q' . $row, $item['sea_remark'] ?? '');
                    $row++;
                }
                
            } elseif ($type === 'rail') {
                $headers = [
                    'A' => 'Станция отправления',
                    'B' => 'Пункт назначения',
                    'C' => 'Тип контейнера',
                    'D' => 'Собственность контейнера',
                    'E' => 'Опасный груз',
                    'F' => 'Охрана',
                    'G' => 'Агент',
                    'H' => 'Profit (₽)',
                    'I' => 'Базовая стоимость 20DC (<24t) (₽)',
                    'J' => 'Базовая стоимость 20DC (24t-28t) (₽)',
                    'K' => 'Базовая стоимость 40HC (28t) (₽)',
                    'L' => 'Стоимость 20DC (<24t) опасный (₽)',
                    'M' => 'Стоимость 20DC (24t-28t) опасный (₽)',
                    'N' => 'Стоимость 40HC (28t) опасный (₽)',
                    'O' => 'Стоимость охраны (₽)',
                    'P' => 'Итого 20DC (<24t) (₽)',
                    'Q' => 'Итого 20DC (24t-28t) (₽)',
                    'R' => 'Итого 40HC (28t) (₽)',
                    'S' => 'Итого 20DC (<24t) опасный (₽)',
                    'T' => 'Итого 20DC (24t-28t) опасный (₽)',
                    'U' => 'Итого 40HC (28t) опасный (₽)'
                ];
                
                // Заполняем заголовки
                foreach ($headers as $col => $header) {
                    $sheet->setCellValue($col . '1', $header);
                    $sheet->getStyle($col . '1')->getFont()->setBold(true);
                }
                
                // Заполняем данные
                $row = 2;
                foreach ($data as $item) {
                    $sheet->setCellValue('A' . $row, $item['rail_origin'] ?? '');
                    $sheet->setCellValue('B' . $row, $item['rail_destination'] ?? '');
                    $sheet->setCellValue('C' . $row, $item['rail_coc'] ?? '');
                    $sheet->setCellValue('D' . $row, $item['rail_container_ownership'] ?? '');
                    $sheet->setCellValue('E' . $row, $item['rail_hazard'] ?? 'Нет');
                    $sheet->setCellValue('F' . $row, $item['rail_security'] ?? 'Нет');
                    $sheet->setCellValue('G' . $row, $item['rail_agent'] ?? '');
                    $sheet->setCellValue('H' . $row, $item['rail_profit'] ?? 0);
                    $sheet->setCellValue('I' . $row, $item['cost_base_20'] ?? 0);
                    $sheet->setCellValue('J' . $row, $item['cost_base_20_28'] ?? 0);
                    $sheet->setCellValue('K' . $row, $item['cost_base_40'] ?? 0);
                    $sheet->setCellValue('L' . $row, $item['cost_base_20_danger'] ?? 0);
                    $sheet->setCellValue('M' . $row, $item['cost_base_20_28_danger'] ?? 0);
                    $sheet->setCellValue('N' . $row, $item['cost_base_40_danger'] ?? 0);
                    $sheet->setCellValue('O' . $row, $item['cost_security'] ?? 0);
                    $sheet->setCellValue('P' . $row, $item['cost_total_20'] ?? 0);
                    $sheet->setCellValue('Q' . $row, $item['cost_total_20_28'] ?? 0);
                    $sheet->setCellValue('R' . $row, $item['cost_total_40'] ?? 0);
                    $sheet->setCellValue('S' . $row, $item['cost_total_20_danger'] ?? 0);
                    $sheet->setCellValue('T' . $row, $item['cost_total_20_28_danger'] ?? 0);
                    $sheet->setCellValue('U' . $row, $item['cost_total_40_danger'] ?? 0);
                    $row++;
                }
                
            } elseif ($type === 'combined') {
                $headers = [
                    'A' => 'Морской порт отправления',
                    'B' => 'Морской порт прибытия',
                    'C' => 'ЖД станция отправления',
                    'D' => 'ЖД станция назначения',
                    'E' => 'DROP OFF LOCATION',
                    'F' => 'Тип контейнера',
                    'G' => 'Собственность контейнера',
                    'H' => 'Опасный груз',
                    'I' => 'Охрана',
                    'J' => 'Стоимость морской части ($)',
                    'K' => 'Стоимость морской части опасный ($)',
                    'L' => 'Стоимость ЖД части (₽)',
                    'M' => 'Стоимость ЖД части опасный (₽)',
                    'N' => 'Общая стоимость ($ + ₽)',
                    'O' => 'Общая стоимость опасный ($ + ₽)',
                    'P' => 'Агент(-ы)',
                    'Q' => 'Комментарий'
                ];
                
                // Заполняем заголовки
                foreach ($headers as $col => $header) {
                    $sheet->setCellValue($col . '1', $header);
                    $sheet->getStyle($col . '1')->getFont()->setBold(true);
                }
                
                // Заполняем данные
                $row = 2;
                foreach ($data as $item) {
                    $sheet->setCellValue('A' . $row, $item['comb_sea_pol'] ?? '');
                    $sheet->setCellValue('B' . $row, $item['comb_sea_pod'] ?? '');
                    $sheet->setCellValue('C' . $row, $item['comb_rail_start'] ?? '');
                    $sheet->setCellValue('D' . $row, $item['comb_rail_dest'] ?? '');
                    $sheet->setCellValue('E' . $row, $item['drop_off_location'] ?? '');
                    $sheet->setCellValue('F' . $row, $item['comb_coc'] ?? '');
                    $sheet->setCellValue('G' . $row, $item['comb_container_ownership'] ?? '');
                    $sheet->setCellValue('H' . $row, $item['comb_hazard'] ?? 'Нет');
                    $sheet->setCellValue('I' . $row, $item['comb_security'] ?? '');
                    $sheet->setCellValue('J' . $row, $item['cost_sea'] ?? 0);
                    $sheet->setCellValue('K' . $row, $item['cost_sea_danger'] ?? ($item['cost_sea'] ?? 0));
                    $sheet->setCellValue('L' . $row, $item['cost_rail'] ?? 0);
                    $sheet->setCellValue('M' . $row, $item['cost_rail_danger'] ?? ($item['cost_rail'] ?? 0));
                    $sheet->setCellValue('N' . $row, $item['cost_total'] ?? 0);
                    $sheet->setCellValue('O' . $row, $item['cost_total_danger'] ?? ($item['cost_total'] ?? 0));
                    $sheet->setCellValue('P' . $row, $item['agent'] ?? '');
                    $sheet->setCellValue('Q' . $row, $item['remark'] ?? '');
                    $row++;
                }
            }
            
            // Авторазмер колонок
            foreach (range('A', $sheet->getHighestColumn()) as $column) {
                $sheet->getColumnDimension($column)->setAutoSize(true);
            }
            
            // Установка границ для всех ячеек
            $lastColumn = $sheet->getHighestColumn();
            $lastRow = $sheet->getHighestRow();
            $styleArray = [
                'borders' => [
                    'allBorders' => [
                        'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                    ],
                ],
            ];
            $sheet->getStyle('A1:' . $lastColumn . $lastRow)->applyFromArray($styleArray);
            
            // Сохраняем файл
            $writer = new Xlsx($spreadsheet);
            
            header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            header('Content-Disposition: attachment;filename="' . $filename . '.xlsx"');
            header('Cache-Control: max-age=0');
            
            $writer->save('php://output');
            exit;
            
        } catch (\Exception $e) {
            header('Content-Type: application/json; charset=utf-8');
            echo json_encode([
                'error' => true,
                'message' => 'Ошибка при экспорте в Excel: ' . $e->getMessage()
            ]);
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
            // Получаем данные морских перевозок с фильтрацией
            $seaPerevozki = self::fetchTransportData(
                28, 
                self::SEA_TRANSPORT_MAP,
                [
                    '=NAME' => $params['sea_pol'] ?? '',
                    'PROPERTY_126' => $params['sea_pod'] ?? '',
                    'PROPERTY_132' => $params['sea_drop_off_location'] ?? '',
                ]
            );
            
            // Если не нашли по точному совпадению DROP_OFF_LOCATION, ищем только по POL и POD
            if (empty($seaPerevozki) && !empty($params['sea_pol']) && !empty($params['sea_pod'])) {
                $seaPerevozki = self::fetchTransportData(
                    28, 
                    self::SEA_TRANSPORT_MAP,
                    [
                        '=NAME' => $params['sea_pol'],
                        'PROPERTY_126' => $params['sea_pod'],
                    ]
                );
            }
            
            // Проверяем, является ли груз опасным
            $isHazard = ($params['sea_hazard'] ?? 'no') === 'yes';
            $security = $params['sea_security'] ?? 'no';
            
            // Обрабатываем каждую найденную запись
            foreach ($seaPerevozki as $value) {
                // Определяем стоимость в зависимости от типа контейнера и собственности
                $containerType = $params['sea_coc'] ?? '';
                $containerOwnership = $params['sea_container_ownership'] ?? 'no';
                $cafPercent = floatval($params['sea_caf'] ?? 0);
                $profit = floatval($params['sea_profit'] ?? 0);
                
                // Получаем базовые стоимости (для обычного и опасного груза)
                $cocCost20 = ceil(floatval($value['COC_20GP'] ?? 0));
                $cocCost40 = ceil(floatval($value['COC_40HC'] ?? 0));
                $socCost20 = ceil(floatval($value['SOC_20GP'] ?? 0));
                $socCost40 = ceil(floatval($value['SOC_40HC'] ?? 0));
                
                // Стоимость опасного груза (если есть)
                $dangerCocCost20 = ceil(floatval($value['OPASNYY_20DC'] ?? 0));
                $dangerCocCost40 = ceil(floatval($value['OPASNYY_40HC'] ?? 0));
                
                // DROP OFF стоимости
                $dropOffCost20 = ceil(floatval($value['DROP_OFF_20GP'] ?? 0));
                $dropOffCost40 = ceil(floatval($value['DROP_OFF_40HC'] ?? 0));
                
                // Стоимость охраны
                $securityCost20 = ceil(floatval($value['OKHRANA_20_FUT'] ?? 0));
                $securityCost40 = ceil(floatval($value['OKHRANA_40_FUT'] ?? 0));
                
                // Определяем стоимость охраны в зависимости от выбора
                $securityCost = 0;
                if ($security === '20') {
                    $securityCost = $securityCost20;
                } elseif ($security === '40') {
                    $securityCost = $securityCost40;
                }
                
                // Если собственность контейнера не выбрана ('no') - показываем оба варианта
                if ($containerOwnership === 'no') {
                    // Расчет для обоих типов контейнеров (20GP и 40HC) если не выбран конкретный
                    $containerTypesToShow = [];
                    if (empty($containerType)) {
                        $containerTypesToShow = ['20GP', '40HC'];
                    } else {
                        $containerTypesToShow = [$containerType];
                    }
                    
                    foreach ($containerTypesToShow as $cType) {
                        $is20GP = ($cType === '20GP');
                        
                        // Базовые стоимости для текущего типа контейнера
                        $dropOffCost = $is20GP ? $dropOffCost20 : $dropOffCost40;
                        $securityCostForType = $is20GP ? $securityCost20 : $securityCost40;
                        $actualSecurityCost = ($security === '20' && $is20GP) || ($security === '40' && !$is20GP) ? $securityCostForType : 0;
                        
                        // Вариант COC - обычный груз
                        $cocCostNormal = $is20GP ? $cocCost20 : $cocCost40;
                        $cocCostDanger = $is20GP ? $dangerCocCost20 : $dangerCocCost40;
                        
                        $cocNettoNormal = $dropOffCost + $cocCostNormal;
                        $cocTotalNormal = ceil(($dropOffCost + $cocNettoNormal) * (1 + $cafPercent / 100) + $profit + $actualSecurityCost);
                        
                        $cocNettoDanger = $dropOffCost + $cocCostDanger;
                        $cocTotalDanger = ceil(($dropOffCost + $cocNettoDanger) * (1 + $cafPercent / 100) + $profit + $actualSecurityCost);
                        
                        // Вариант COC - опасный груз
                        $cocCostDanger = $is20GP ? $dangerCocCost20 : $dangerCocCost40;
                        
                        // Вариант SOC - обычный груз
                        $socCostNormal = $is20GP ? $socCost20 : $socCost40;
                        $socNettoNormal = $dropOffCost + $socCostNormal;
                        $socTotalNormal = ceil(($dropOffCost + $socNettoNormal) * (1 + $cafPercent / 100) + $profit + $actualSecurityCost);
                        
                        // Вариант SOC - опасный груз (используем те же стоимости, что и COC для опасного груза)
                        $socNettoDanger = $dropOffCost + $cocCostDanger; // Для SOC опасного используем стоимости опасного COC
                        $socTotalDanger = ceil(($dropOffCost + $socNettoDanger) * (1 + $cafPercent / 100) + $profit + $actualSecurityCost);
                        
                        $result[] = [
                            'sea_pol' => $value['POL'] ?? '',
                            'sea_pod' => $value['POD'] ?? '',
                            'sea_drop_off_location' => $value['DROP_OFF_LOCATION'] ?? '',
                            'sea_coc' => $cType,
                            'sea_container_ownership' => 'COC: контейнер линии',
                            'sea_agent' => $value['AGENT'] ?? '',
                            'sea_remark' => $value['REMARK'] ?? '',
                            'sea_hazard' => 'Нет',
                            'sea_security' => $security === 'no' ? 'Нет' : ($security === '20' ? '20 фут' : '40 фут'),
                            'sea_caf_percent' => $cafPercent,
                            'sea_profit' => $profit,
                            'cost_container' => $cocCostNormal,
                            'cost_container_danger' => $cocCostDanger,
                            'cost_drop_off' => $dropOffCost,
                            'cost_security' => $actualSecurityCost,
                            'cost_netto' => $cocNettoNormal,
                            'cost_total' => $cocTotalNormal,
                            'cost_total_danger' => $cocTotalDanger,
                            'calculation_formula' => "($dropOffCost + $cocNettoNormal) × (1 + $cafPercent/100) + $profit + $actualSecurityCost = " . number_format($cocTotalNormal, 2) . " $",
                            'calculation_formula_danger' => "($dropOffCost + $cocNettoDanger) × (1 + $cafPercent/100) + $profit + $actualSecurityCost = " . number_format($cocTotalDanger, 2) . " $",
                            'show_both_ownership' => true,
                            'show_both_hazard' => true,
                            'soc_option' => [
                                'sea_container_ownership' => 'SOC: контейнер агента',
                                'cost_container' => $socCostNormal,
                                'cost_container_danger' => $cocCostDanger,
                                'cost_netto' => $socNettoNormal,
                                'cost_netto_danger' => $socNettoDanger,
                                'cost_total' => $socTotalNormal,
                                'cost_total_danger' => $socTotalDanger,
                                'calculation_formula' => "($dropOffCost + $socNettoNormal) × (1 + $cafPercent/100) + $profit + $actualSecurityCost = " . number_format($socTotalNormal, 2) . " $",
                                'calculation_formula_danger' => "($dropOffCost + $socNettoDanger) × (1 + $cafPercent/100) + $profit + $actualSecurityCost = " . number_format($socTotalDanger, 2) . " $"
                            ]
                        ];
                        
                        // Вариант для опасного груза COC
                        $result[] = [
                            'sea_pol' => $value['POL'] ?? '',
                            'sea_pod' => $value['POD'] ?? '',
                            'sea_drop_off_location' => $value['DROP_OFF_LOCATION'] ?? '',
                            'sea_coc' => $cType,
                            'sea_container_ownership' => 'COC: контейнер линии (опасный)',
                            'sea_agent' => $value['AGENT'] ?? '',
                            'sea_remark' => $value['REMARK'] ?? '',
                            'sea_hazard' => 'Да',
                            'sea_security' => $security === 'no' ? 'Нет' : ($security === '20' ? '20 фут' : '40 фут'),
                            'sea_caf_percent' => $cafPercent,
                            'sea_profit' => $profit,
                            'cost_container' => $cocCostDanger,
                            'cost_drop_off' => $dropOffCost,
                            'cost_security' => $actualSecurityCost,
                            'cost_netto' => $cocNettoDanger,
                            'cost_total' => $cocTotalDanger,
                            'calculation_formula' => "($dropOffCost + $cocNettoDanger) × (1 + $cafPercent/100) + $profit + $actualSecurityCost = " . number_format($cocTotalDanger, 2) . " $",
                            'show_both_ownership' => true,
                            'show_both_hazard' => true,
                            'normal_option' => [
                                'sea_container_ownership' => 'COC: контейнер линии (обычный)',
                                'cost_container' => $cocCostNormal,
                                'cost_netto' => $cocNettoNormal,
                                'cost_total' => $cocTotalNormal,
                                'calculation_formula' => "($dropOffCost + $cocNettoNormal) × (1 + $cafPercent/100) + $profit + $actualSecurityCost = " . number_format($cocTotalNormal, 2) . " $"
                            ]
                        ];
                        
                        // Вариант SOC - обычный груз
                        $result[] = [
                            'sea_pol' => $value['POL'] ?? '',
                            'sea_pod' => $value['POD'] ?? '',
                            'sea_drop_off_location' => $value['DROP_OFF_LOCATION'] ?? '',
                            'sea_coc' => $cType,
                            'sea_container_ownership' => 'SOC: контейнер агента',
                            'sea_agent' => $value['AGENT'] ?? '',
                            'sea_remark' => $value['REMARK'] ?? '',
                            'sea_hazard' => 'Нет',
                            'sea_security' => $security === 'no' ? 'Нет' : ($security === '20' ? '20 фут' : '40 фут'),
                            'sea_caf_percent' => $cafPercent,
                            'sea_profit' => $profit,
                            'cost_container' => $socCostNormal,
                            'cost_container_danger' => $cocCostDanger,
                            'cost_drop_off' => $dropOffCost,
                            'cost_security' => $actualSecurityCost,
                            'cost_netto' => $socNettoNormal,
                            'cost_total' => $socTotalNormal,
                            'cost_total_danger' => $socTotalDanger,
                            'calculation_formula' => "($dropOffCost + $socNettoNormal) × (1 + $cafPercent/100) + $profit + $actualSecurityCost = " . number_format($socTotalNormal, 2) . " $",
                            'calculation_formula_danger' => "($dropOffCost + $socNettoDanger) × (1 + $cafPercent/100) + $profit + $actualSecurityCost = " . number_format($socTotalDanger, 2) . " $",
                            'show_both_ownership' => true,
                            'show_both_hazard' => true,
                            'coc_option' => [
                                'sea_container_ownership' => 'COC: контейнер линии',
                                'cost_container' => $cocCostNormal,
                                'cost_container_danger' => $cocCostDanger,
                                'cost_netto' => $cocNettoNormal,
                                'cost_netto_danger' => $cocNettoDanger,
                                'cost_total' => $cocTotalNormal,
                                'cost_total_danger' => $cocTotalDanger,
                                'calculation_formula' => "($dropOffCost + $cocNettoNormal) × (1 + $cafPercent/100) + $profit + $actualSecurityCost = " . number_format($cocTotalNormal, 2) . " $",
                                'calculation_formula_danger' => "($dropOffCost + $cocNettoDanger) × (1 + $cafPercent/100) + $profit + $actualSecurityCost = " . number_format($cocTotalDanger, 2) . " $"
                            ]
                        ];
                        
                        // Вариант SOC - опасный груз
                        $result[] = [
                            'sea_pol' => $value['POL'] ?? '',
                            'sea_pod' => $value['POD'] ?? '',
                            'sea_drop_off_location' => $value['DROP_OFF_LOCATION'] ?? '',
                            'sea_coc' => $cType,
                            'sea_container_ownership' => 'SOC: контейнер агента (опасный)',
                            'sea_agent' => $value['AGENT'] ?? '',
                            'sea_remark' => $value['REMARK'] ?? '',
                            'sea_hazard' => 'Да',
                            'sea_security' => $security === 'no' ? 'Нет' : ($security === '20' ? '20 фут' : '40 фут'),
                            'sea_caf_percent' => $cafPercent,
                            'sea_profit' => $profit,
                            'cost_container' => $cocCostDanger,
                            'cost_drop_off' => $dropOffCost,
                            'cost_security' => $actualSecurityCost,
                            'cost_netto' => $socNettoDanger,
                            'cost_total' => $socTotalDanger,
                            'calculation_formula' => "($dropOffCost + $socNettoDanger) × (1 + $cafPercent/100) + $profit + $actualSecurityCost = " . number_format($socTotalDanger, 2) . " $",
                            'show_both_ownership' => true,
                            'show_both_hazard' => true,
                            'normal_option' => [
                                'sea_container_ownership' => 'SOC: контейнер агента (обычный)',
                                'cost_container' => $socCostNormal,
                                'cost_netto' => $socNettoNormal,
                                'cost_total' => $socTotalNormal,
                                'calculation_formula' => "($dropOffCost + $socNettoNormal) × (1 + $cafPercent/100) + $profit + $actualSecurityCost = " . number_format($socTotalNormal, 2) . " $"
                            ]
                        ];
                    }
                } else {
                    // Если выбрана конкретная собственность контейнера и тип опасности
                    $containerTypesToShow = [];
                    if (empty($containerType)) {
                        $containerTypesToShow = ['20GP', '40HC'];
                    } else {
                        $containerTypesToShow = [$containerType];
                    }
                    
                    foreach ($containerTypesToShow as $cType) {
                        $is20GP = ($cType === '20GP');
                        
                        // Базовые стоимости
                        $dropOffCost = $is20GP ? $dropOffCost20 : $dropOffCost40;
                        $securityCostForType = $is20GP ? $securityCost20 : $securityCost40;
                        $actualSecurityCost = ($security === '20' && $is20GP) || ($security === '40' && !$is20GP) ? $securityCostForType : 0;
                        
                        // Определяем стоимости в зависимости от типа собственности и опасности
                        if ($isHazard) {
                            // Опасный груз
                            $containerCost = $containerOwnership === 'soc' 
                                ? ($is20GP ? $dangerCocCost20 : $dangerCocCost40)
                                : ($is20GP ? $dangerCocCost20 : $dangerCocCost40);
                        } else {
                            // Обычный груз
                            $containerCost = $containerOwnership === 'soc' 
                                ? ($is20GP ? $socCost20 : $socCost40)
                                : ($is20GP ? $cocCost20 : $cocCost40);
                        }
                        
                        $netto = $dropOffCost + $containerCost;
                        $totalCost = ceil(($dropOffCost + $netto) * (1 + $cafPercent / 100) + $profit + $actualSecurityCost);
                        
                        $displayContainerType = $containerOwnership === 'coc' 
                            ? 'COC: контейнер линии' . ($isHazard ? ' (опасный)' : '') 
                            : 'SOC: контейнер агента' . ($isHazard ? ' (опасный)' : '');
                        
                        $result[] = [
                            'sea_pol' => $value['POL'] ?? '',
                            'sea_pod' => $value['POD'] ?? '',
                            'sea_drop_off_location' => $value['DROP_OFF_LOCATION'] ?? '',
                            'sea_coc' => $cType,
                            'sea_container_ownership' => $displayContainerType,
                            'sea_agent' => $value['AGENT'] ?? '',
                            'sea_remark' => $value['REMARK'] ?? '',
                            'sea_hazard' => $isHazard ? 'Да' : 'Нет',
                            'sea_security' => $security === 'no' ? 'Нет' : ($security === '20' ? '20 фут' : '40 фут'),
                            'sea_caf_percent' => $cafPercent,
                            'sea_profit' => $profit,
                            'cost_container' => $containerCost,
                            'cost_drop_off' => $dropOffCost,
                            'cost_security' => $actualSecurityCost,
                            'cost_netto' => $netto,
                            'cost_total' => $totalCost,
                            'calculation_formula' => "($dropOffCost + $netto) × (1 + $cafPercent/100) + $profit + $actualSecurityCost = " . number_format($totalCost, 2) . " $",
                            'show_both_ownership' => false,
                            'show_both_hazard' => false
                        ];
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
                30, 
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
                $containerOwnership = $params['rail_container_ownership'] ?? 'no';
                
                // Базовая стоимость для всех типов контейнеров
                $baseCost20 = ceil($isHazard ? floatval($value['OPASNYY_20DC_24'] ?? 0) : floatval($value['DC20_24'] ?? 0));
                $baseCost20_28 = ceil($isHazard ? floatval($value['OPASNYY_DC20_24T_28T'] ?? 0) : floatval($value['DC20_24T_28T'] ?? 0));
                $baseCost40 = ceil($isHazard ? floatval($value['OPASNYY_HC40_28T'] ?? 0) : floatval($value['HC40_28T'] ?? 0));
                
                // Стоимость охраны
                $securityCost = 0;
                if ($security === '20') {
                    $securityCost = ceil(floatval($value['OKHRANA_20_FUT'] ?? 0));
                } elseif ($security === '40') {
                    $securityCost = ceil(floatval($value['OKHRANA_40_FUT'] ?? 0));
                }
                
                // Общая стоимость для каждого типа контейнера
                $totalCost20 = ceil($baseCost20 + $securityCost + $profit);
                $totalCost20_28 = ceil($baseCost20_28 + $securityCost + $profit);
                $totalCost40 = ceil($baseCost40 + $securityCost + $profit);
                
                // Если собственность контейнера не выбрана - показываем оба варианта
                if ($containerOwnership === 'no') {
                    // Вариант COC
                    // В методе getRailPerevozki() обновляем вывод данных:
                    $result[] = [
                        'rail_origin' => $value['POL'] ?? '',
                        'rail_destination' => $value['POD'] ?? '',
                        'rail_coc' => $cocType,
                        'rail_container_ownership' => 'COC: контейнер линии',
                        'rail_agent' => $value['AGENT'] ?? '',
                        'rail_hazard' => 'Нет', // Добавляем явное указание
                        'rail_security' => $security === 'no' ? 'Нет' : ($security === '20' ? '20 фут' : '40 фут'),
                        'rail_profit' => $profit,
                        'cost_base_20' => $baseCost20,
                        'cost_base_20_28' => $baseCost20_28,
                        'cost_base_40' => $baseCost40,
                        'cost_base_20_danger' => $baseCost20Danger,
                        'cost_base_20_28_danger' => $baseCost20_28Danger,
                        'cost_base_40_danger' => $baseCost40Danger,
                        'cost_security' => $securityCost,
                        'cost_total_20' => $totalCost20,
                        'cost_total_20_28' => $totalCost20_28,
                        'cost_total_40' => $totalCost40,
                        'cost_total_20_danger' => $totalCost20Danger,
                        'cost_total_20_28_danger' => $totalCost20_28Danger,
                        'cost_total_40_danger' => $totalCost40Danger,
                        'calculation_formula_20' => "$baseCost20 (базовая) + $securityCost (охрана) + $profit (прибыль) = $totalCost20 ₽",
                        'calculation_formula_20_28' => "$baseCost20_28 (базовая) + $securityCost (охрана) + $profit (прибыль) = $totalCost20_28 ₽",
                        'calculation_formula_40' => "$baseCost40 (базовая) + $securityCost (охрана) + $profit (прибыль) = $totalCost40 ₽",
                        'calculation_formula_20_danger' => "$baseCost20Danger (базовая) + $securityCost (охрана) + $profit (прибыль) = $totalCost20Danger ₽",
                        'calculation_formula_20_28_danger' => "$baseCost20_28Danger (базовая) + $securityCost (охрана) + $profit (прибыль) = $totalCost20_28Danger ₽",
                        'calculation_formula_40_danger' => "$baseCost40Danger (базовая) + $securityCost (охрана) + $profit (прибыль) = $totalCost40Danger ₽",
                        'show_both_ownership' => true,
                        'show_both_hazard' => true,
                        'soc_option' => [
                            'rail_container_ownership' => 'SOC: контейнер агента',
                            'cost_total_20' => $totalCost20,
                            'cost_total_20_28' => $totalCost20_28,
                            'cost_total_40' => $totalCost40,
                            'cost_total_20_danger' => $totalCost20Danger,
                            'cost_total_20_28_danger' => $totalCost20_28Danger,
                            'cost_total_40_danger' => $totalCost40Danger,
                            'calculation_formula_20' => "$baseCost20 (базовая) + $securityCost (охрана) + $profit (прибыль) = $totalCost20 ₽",
                            'calculation_formula_20_28' => "$baseCost20_28 (базовая) + $securityCost (охрана) + $profit (прибыль) = $totalCost20_28 ₽",
                            'calculation_formula_40' => "$baseCost40 (базовая) + $securityCost (охрана) + $profit (прибыль) = $totalCost40 ₽",
                            'calculation_formula_20_danger' => "$baseCost20Danger (базовая) + $securityCost (охрана) + $profit (прибыль) = $totalCost20Danger ₽",
                            'calculation_formula_20_28_danger' => "$baseCost20_28Danger (базовая) + $securityCost (охрана) + $profit (прибыль) = $totalCost20_28Danger ₽",
                            'calculation_formula_40_danger' => "$baseCost40Danger (базовая) + $securityCost (охрана) + $profit (прибыль) = $totalCost40Danger ₽",
                        ]
                    ];

                    // Также добавляем вариант для опасного груза
                    $result[] = [
                        'rail_origin' => $value['POL'] ?? '',
                        'rail_destination' => $value['POD'] ?? '',
                        'rail_coc' => $cocType,
                        'rail_container_ownership' => 'COC: контейнер линии (опасный)',
                        'rail_agent' => $value['AGENT'] ?? '',
                        'rail_hazard' => 'Да',
                        'rail_security' => $security === 'no' ? 'Нет' : ($security === '20' ? '20 фут' : '40 фут'),
                        'rail_profit' => $profit,
                        'cost_base_20' => $baseCost20Danger,
                        'cost_base_20_28' => $baseCost20_28Danger,
                        'cost_base_40' => $baseCost40Danger,
                        'cost_security' => $securityCost,
                        'cost_total_20' => $totalCost20Danger,
                        'cost_total_20_28' => $totalCost20_28Danger,
                        'cost_total_40' => $totalCost40Danger,
                        'calculation_formula_20' => "$baseCost20Danger (базовая) + $securityCost (охрана) + $profit (прибыль) = $totalCost20Danger ₽",
                        'calculation_formula_20_28' => "$baseCost20_28Danger (базовая) + $securityCost (охрана) + $profit (прибыль) = $totalCost20_28Danger ₽",
                        'calculation_formula_40' => "$baseCost40Danger (базовая) + $securityCost (охрана) + $profit (прибыль) = $totalCost40Danger ₽",
                        'show_both_ownership' => true,
                        'show_both_hazard' => true,
                        'normal_option' => [
                            'rail_container_ownership' => 'COC: контейнер линии (обычный)',
                            'cost_total_20' => $totalCost20,
                            'cost_total_20_28' => $totalCost20_28,
                            'cost_total_40' => $totalCost40,
                            'calculation_formula_20' => "$baseCost20 (базовая) + $securityCost (охрана) + $profit (прибыль) = $totalCost20 ₽",
                            'calculation_formula_20_28' => "$baseCost20_28 (базовая) + $securityCost (охрана) + $profit (прибыль) = $totalCost20_28 ₽",
                            'calculation_formula_40' => "$baseCost40 (базовая) + $securityCost (охрана) + $profit (прибыль) = $totalCost40 ₽",
                        ]
                    ];
                } else {
                    // Если выбрана конкретная собственность контейнера
                    $containerType = $containerOwnership === 'coc' 
                        ? 'COC: контейнер линии' 
                        : 'SOC: контейнер агента';
                    
                    $result[] = [
                        'rail_origin' => $value['POL'] ?? '',
                        'rail_destination' => $value['POD'] ?? '',
                        'rail_coc' => $cocType,
                        'rail_container_ownership' => $containerType,
                        'rail_agent' => $value['AGENT'] ?? '',
                        'rail_hazard' => $isHazard ? 'Да' : 'Нет',
                        'rail_security' => $security === 'no' ? 'Нет' : ($security === '20' ? '20 фут' : '40 фут'),
                        'rail_profit' => $profit,
                        'cost_base_20' => $baseCost20,
                        'cost_base_20_28' => $baseCost20_28,
                        'cost_base_40' => $baseCost40,
                        'cost_security' => $securityCost,
                        'cost_total_20' => $totalCost20,
                        'cost_total_20_28' => $totalCost20_28,
                        'cost_total_40' => $totalCost40,
                        'calculation_formula_20' => "$baseCost20 (базовая) + $securityCost (охрана) + $profit (прибыль) = $totalCost20 ₽",
                        'calculation_formula_20_28' => "$baseCost20_28 (базовая) + $securityCost (охрана) + $profit (прибыль) = $totalCost20_28 ₽",
                        'calculation_formula_40' => "$baseCost40 (базовая) + $securityCost (охрана) + $profit (прибыль) = $totalCost40 ₽",
                        'show_both_ownership' => false
                    ];
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
     * Получаем комбинированные маршруты
     *
     * @return [type]
     * 
     */
    public function getCombPerevozki() {
        $result = [];
        $params = $_POST;
        header('Content-Type: application/json; charset=utf-8');
        
        try {
            // Получаем морские перевозки с учетом DROP OFF
            $seaFilter = [
                '=NAME' => $params['comb_sea_pol'] ?? '',
            ];
            
            // Если указан DROP OFF, добавляем его в фильтр
            if (!empty($params['comb_drop_off'])) {
                $seaFilter['=PROPERTY_132'] = $params['comb_drop_off'];
            }

            if (!empty($params['comb_transshipment_port'])) {
                $seaFilter['=PROPERTY_126'] = $params['comb_transshipment_port'];
            }

            $seaPerevozki = self::fetchTransportData(
                28, 
                self::SEA_TRANSPORT_MAP,
                $seaFilter
            );
            
            // Получаем ж/д перевозки
            $combFilter = [
                '=PROPERTY_188' => $params['comb_rail_dest'] ?? '',
            ];

            $combPerevozki = self::fetchTransportData(
                32,
                self::COMB_TRANSPORT_MAP,
                $combFilter
            );

            // Фильтруем станции отправления
            $combStartStations = array_unique(array_column($combPerevozki, 'STANTSIYA_OTPRAVLENIYA'));
            
            // Получаем ж/д перевозки
            $zhdPerevozki = self::fetchTransportData(
                30, 
                self::ZHD_TRANSPORT_MAP,
                [
                    '=NAME' => $combStartStations,
                    '=PROPERTY_142' => $params['comb_rail_dest'],
                ]
            );
            
            // Параметры опасности и охраны
            $isHazard = ($params['comb_hazard'] ?? 'no') === 'yes';
            $security = $params['comb_security'] ?? 'no';
            
            // Определяем собственность контейнера
            $containerOwnership = $params['comb_container_ownership'] ?? 'no';
            
            // Если собственность не выбрана ('no') - показываем оба варианта
            if ($containerOwnership === 'no') {
                // Вариант COC - обычный груз
                foreach ($seaPerevozki as $value) {
                    // Расчет морской части для COC
                    $coc = $params['comb_coc'] != '40HC (28t)' ? ceil(floatval($value['COC_20GP'] ?? 0)) : ceil(floatval($value['COC_40HC'] ?? 0));
                    $cocDanger = $params['comb_coc'] != '40HC (28t)' ? ceil(floatval($value['OPASNYY_20DC'] ?? 0)) : ceil(floatval($value['OPASNYY_40HC'] ?? 0));
                    
                    $dropOff = $params['comb_coc'] != '40HC (28t)' ? ceil(floatval($value['DROP_OFF_20GP'] ?? 0)) : ceil(floatval($value['DROP_OFF_40HC'] ?? 0));
                    
                    // Стоимость охраны для морской части
                    $securityCostSea = 0;
                    if ($security === '20' && $params['comb_coc'] != '40HC (28t)') {
                        $securityCostSea = ceil(floatval($value['OKHRANA_20_FUT'] ?? 0));
                    } elseif ($security === '40' && $params['comb_coc'] == '40HC (28t)') {
                        $securityCostSea = ceil(floatval($value['OKHRANA_40_FUT'] ?? 0));
                    }
                    
                    $netto = ceil($dropOff + $coc);
                    $nettoDanger = ceil($dropOff + $cocDanger);
                    $costSea = ceil(($dropOff + $netto) * (floatval($value['CAF_KONVERT'] ?? 0) / 100 + 1) + $securityCostSea);
                    $costSeaDanger = ceil(($dropOff + $nettoDanger) * (floatval($value['CAF_KONVERT'] ?? 0) / 100 + 1) + $securityCostSea);

                    foreach ($zhdPerevozki as $v) {
                        if ($params['comb_coc'] == '40HC (28t)') {
                            $cost = ceil(floatval($v['HC40_28T'] ?? 0));
                            $costDanger = ceil(floatval($v['OPASNYY_HC40_28T'] ?? 0));
                        } elseif ($params['comb_coc'] == '20DC (24t-28t)') {
                            $cost = ceil(floatval($v['DC20_24T_28T'] ?? 0));
                            $costDanger = ceil(floatval($v['OPASNYY_DC20_24T_28T'] ?? 0));
                        } else {
                            $cost = ceil(floatval($v['DC20_24'] ?? 0));
                            $costDanger = ceil(floatval($v['OPASNYY_20DC_24'] ?? 0));
                        }
                        
                        // Стоимость охраны для ЖД части
                        $protect = $params['comb_security'] == '40 фут' ? ceil(floatval($v['OKHRANA_40_FUT'] ?? 0)) : ($params['comb_security'] == '20 фут' ? ceil(floatval($v['OKHRANA_20_FUT'] ?? 0)) : 0);
                        
                        $costZhd = ceil($cost + $protect);
                        $costZhdDanger = ceil($costDanger + $protect);
                        
                        // Общая стоимость (море + ЖД)
                        $totalCost = $costSea + $costZhd;
                        $totalCostDanger = $costSeaDanger + $costZhdDanger;
                        
                        $result[] = [
                            'comb_sea_pol' => $value['POL'] ?? '',
                            'comb_sea_pod' => $value['POD'] ?? '',
                            'comb_rail_start' => $v['POL'] ?? '',
                            'comb_rail_dest' => $v['POD'] ?? '',
                            'comb_security' => $params['comb_security'] ?? '',
                            'comb_hazard' => 'Нет',
                            'comb_coc' => $params['comb_coc'] ?? '',
                            'comb_container_ownership' => 'COC: контейнер линии',
                            'cost_sea' => $costSea,
                            'cost_sea_danger' => $costSeaDanger,
                            'cost_rail' => $costZhd,
                            'cost_rail_danger' => $costZhdDanger,
                            'cost_total' => $totalCost,
                            'cost_total_danger' => $totalCostDanger,
                            'drop_off_location' => $value['DROP_OFF_LOCATION'] ?? '',
                            'agent' => implode(', ', array_filter([$value['AGENT'] ?? '', $v['AGENT'] ?? ''])),
                            'show_both_ownership' => true,
                            'show_both_hazard' => true,
                            'soc_option' => [
                                'comb_container_ownership' => 'SOC: контейнер агента',
                                'cost_sea' => $costSea,
                                'cost_sea_danger' => $costSeaDanger,
                                'cost_rail' => $costZhd,
                                'cost_rail_danger' => $costZhdDanger,
                                'cost_total' => $totalCost,
                                'cost_total_danger' => $totalCostDanger,
                            ],
                            'danger_option' => [
                                'comb_container_ownership' => 'COC: контейнер линии (опасный)',
                                'cost_sea' => $costSeaDanger,
                                'cost_rail' => $costZhdDanger,
                                'cost_total' => $totalCostDanger,
                            ]
                        ];
                        
                        // Вариант COC - опасный груз
                        $result[] = [
                            'comb_sea_pol' => $value['POL'] ?? '',
                            'comb_sea_pod' => $value['POD'] ?? '',
                            'comb_rail_start' => $v['POL'] ?? '',
                            'comb_rail_dest' => $v['POD'] ?? '',
                            'comb_security' => $params['comb_security'] ?? '',
                            'comb_hazard' => 'Да',
                            'comb_coc' => $params['comb_coc'] ?? '',
                            'comb_container_ownership' => 'COC: контейнер линии (опасный)',
                            'cost_sea' => $costSeaDanger,
                            'cost_rail' => $costZhdDanger,
                            'cost_total' => $totalCostDanger,
                            'drop_off_location' => $value['DROP_OFF_LOCATION'] ?? '',
                            'agent' => implode(', ', array_filter([$value['AGENT'] ?? '', $v['AGENT'] ?? ''])),
                            'show_both_ownership' => true,
                            'show_both_hazard' => true,
                            'normal_option' => [
                                'comb_container_ownership' => 'COC: контейнер линии (обычный)',
                                'cost_sea' => $costSea,
                                'cost_rail' => $costZhd,
                                'cost_total' => $totalCost,
                            ]
                        ];
                        
                        // Вариант SOC - обычный груз
                        $soc = $params['comb_coc'] != '40HC (28t)' ? ceil(floatval($value['SOC_20GP'] ?? 0)) : ceil(floatval($value['SOC_40HC'] ?? 0));
                        $nettoSoc = ceil($dropOff + $soc);
                        $costSeaSoc = ceil(($dropOff + $nettoSoc) * (floatval($value['CAF_KONVERT'] ?? 0) / 100 + 1) + $securityCostSea);
                        
                        $result[] = [
                            'comb_sea_pol' => $value['POL'] ?? '',
                            'comb_sea_pod' => $value['POD'] ?? '',
                            'comb_rail_start' => $v['POL'] ?? '',
                            'comb_rail_dest' => $v['POD'] ?? '',
                            'comb_security' => $params['comb_security'] ?? '',
                            'comb_hazard' => 'Нет',
                            'comb_coc' => $params['comb_coc'] ?? '',
                            'comb_container_ownership' => 'SOC: контейнер агента',
                            'cost_sea' => $costSeaSoc,
                            'cost_sea_danger' => $costSeaDanger,
                            'cost_rail' => $costZhd,
                            'cost_rail_danger' => $costZhdDanger,
                            'cost_total' => $costSeaSoc + $costZhd,
                            'cost_total_danger' => $costSeaDanger + $costZhdDanger,
                            'drop_off_location' => $value['DROP_OFF_LOCATION'] ?? '',
                            'agent' => implode(', ', array_filter([$value['AGENT'] ?? '', $v['AGENT'] ?? ''])),
                            'show_both_ownership' => true,
                            'show_both_hazard' => true,
                            'coc_option' => [
                                'comb_container_ownership' => 'COC: контейнер линии',
                                'cost_sea' => $costSea,
                                'cost_sea_danger' => $costSeaDanger,
                                'cost_rail' => $costZhd,
                                'cost_rail_danger' => $costZhdDanger,
                                'cost_total' => $totalCost,
                                'cost_total_danger' => $totalCostDanger,
                            ],
                            'danger_option' => [
                                'comb_container_ownership' => 'SOC: контейнер агента (опасный)',
                                'cost_sea' => $costSeaDanger,
                                'cost_rail' => $costZhdDanger,
                                'cost_total' => $totalCostDanger,
                            ]
                        ];
                        
                        // Вариант SOC - опасный груз
                        $result[] = [
                            'comb_sea_pol' => $value['POL'] ?? '',
                            'comb_sea_pod' => $value['POD'] ?? '',
                            'comb_rail_start' => $v['POL'] ?? '',
                            'comb_rail_dest' => $v['POD'] ?? '',
                            'comb_security' => $params['comb_security'] ?? '',
                            'comb_hazard' => 'Да',
                            'comb_coc' => $params['comb_coc'] ?? '',
                            'comb_container_ownership' => 'SOC: контейнер агента (опасный)',
                            'cost_sea' => $costSeaDanger,
                            'cost_rail' => $costZhdDanger,
                            'cost_total' => $totalCostDanger,
                            'drop_off_location' => $value['DROP_OFF_LOCATION'] ?? '',
                            'agent' => implode(', ', array_filter([$value['AGENT'] ?? '', $v['AGENT'] ?? ''])),
                            'show_both_ownership' => true,
                            'show_both_hazard' => true,
                            'normal_option' => [
                                'comb_container_ownership' => 'SOC: контейнер агента (обычный)',
                                'cost_sea' => $costSeaSoc,
                                'cost_rail' => $costZhd,
                                'cost_total' => $costSeaSoc + $costZhd,
                            ]
                        ];
                    }
                }
            } else {
                // Если выбрана конкретная собственность контейнера и опасность
                $selectedOwnership = $params['comb_container_ownership'] ?? '';
                $isHazard = ($params['comb_hazard'] ?? 'no') === 'yes';
                
                foreach ($seaPerevozki as $value) {
                    // Определяем стоимости в зависимости от типа собственности
                    if ($selectedOwnership === 'coc') {
                        $containerCost = $params['comb_coc'] != '40HC (28t)' 
                            ? ceil(floatval($value['COC_20GP'] ?? 0)) 
                            : ceil(floatval($value['COC_40HC'] ?? 0));
                        $containerCostDanger = $params['comb_coc'] != '40HC (28t)' 
                            ? ceil(floatval($value['OPASNYY_20DC'] ?? 0)) 
                            : ceil(floatval($value['OPASNYY_40HC'] ?? 0));
                    } else {
                        $containerCost = $params['comb_coc'] != '40HC (28t)' 
                            ? ceil(floatval($value['SOC_20GP'] ?? 0)) 
                            : ceil(floatval($value['SOC_40HC'] ?? 0));
                        $containerCostDanger = $params['comb_coc'] != '40HC (28t)' 
                            ? ceil(floatval($value['OPASNYY_20DC'] ?? 0)) 
                            : ceil(floatval($value['OPASNYY_40HC'] ?? 0));
                    }
                    
                    $dropOff = $params['comb_coc'] != '40HC (28t)' 
                        ? ceil(floatval($value['DROP_OFF_20GP'] ?? 0)) 
                        : ceil(floatval($value['DROP_OFF_40HC'] ?? 0));
                    
                    // Стоимость охраны для морской части
                    $securityCostSea = 0;
                    if ($security === '20' && $params['comb_coc'] != '40HC (28t)') {
                        $securityCostSea = ceil(floatval($value['OKHRANA_20_FUT'] ?? 0));
                    } elseif ($security === '40' && $params['comb_coc'] == '40HC (28t)') {
                        $securityCostSea = ceil(floatval($value['OKHRANA_40_FUT'] ?? 0));
                    }
                    
                    // Выбираем стоимость в зависимости от опасности
                    $finalContainerCost = $isHazard ? $containerCostDanger : $containerCost;
                    $netto = ceil($dropOff + $finalContainerCost);
                    $costSea = ceil(($dropOff + $netto) * (floatval($value['CAF_KONVERT'] ?? 0) / 100 + 1) + $securityCostSea);

                    foreach ($zhdPerevozki as $v) {
                        if ($params['comb_coc'] == '40HC (28t)') {
                            $cost = ceil(floatval($v['HC40_28T'] ?? 0));
                            $costDanger = ceil(floatval($v['OPASNYY_HC40_28T'] ?? 0));
                        } elseif ($params['comb_coc'] == '20DC (24t-28t)') {
                            $cost = ceil(floatval($v['DC20_24T_28T'] ?? 0));
                            $costDanger = ceil(floatval($v['OPASNYY_DC20_24T_28T'] ?? 0));
                        } else {
                            $cost = ceil(floatval($v['DC20_24'] ?? 0));
                            $costDanger = ceil(floatval($v['OPASNYY_20DC_24'] ?? 0));
                        }
                        
                        // Стоимость охраны для ЖД части
                        $protect = $params['comb_security'] == '40 фут' 
                            ? ceil(floatval($v['OKHRANA_40_FUT'] ?? 0)) 
                            : ($params['comb_security'] == '20 фут' 
                                ? ceil(floatval($v['OKHRANA_20_FUT'] ?? 0)) 
                                : 0);
                        
                        $finalCostZhd = $isHazard ? ceil($costDanger + $protect) : ceil($cost + $protect);
                        
                        // Общая стоимость
                        $totalCost = $costSea + $finalCostZhd;
                        
                        $displayContainerType = $selectedOwnership === 'coc' 
                            ? 'COC: контейнер линии' . ($isHazard ? ' (опасный)' : '') 
                            : 'SOC: контейнер агента' . ($isHazard ? ' (опасный)' : '');
                        
                        $result[] = [
                            'comb_sea_pol' => $value['POL'] ?? '',
                            'comb_sea_pod' => $value['POD'] ?? '',
                            'comb_rail_start' => $v['POL'] ?? '',
                            'comb_rail_dest' => $v['POD'] ?? '',
                            'comb_security' => $params['comb_security'] ?? '',
                            'comb_hazard' => $isHazard ? 'Да' : 'Нет',
                            'comb_coc' => $params['comb_coc'] ?? '',
                            'comb_container_ownership' => $displayContainerType,
                            'cost_sea' => $costSea,
                            'cost_rail' => $finalCostZhd,
                            'cost_total' => $totalCost,
                            'drop_off_location' => $value['DROP_OFF_LOCATION'] ?? '',
                            'agent' => implode(', ', array_filter([$value['AGENT'] ?? '', $v['AGENT'] ?? ''])),
                            'show_both_ownership' => false,
                            'show_both_hazard' => false
                        ];
                    }
                }
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
                30, 
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
                    $containerType = 'COC: контейнер линии';
                } elseif ($containerOwnership === 'soc') {
                    $containerType = 'SOC: контейнер агента';
                }
                
                $result[] = [
                    'rail_origin' => $value['POL'] ?? '',
                    'rail_destination' => $value['POD'] ?? '',
                    'rail_coc' => $cocType,
                    'rail_container_ownership' => $containerType,
                    'rail_agent' => $value['AGENT'] ?? '',
                    'rail_hazard' => $isHazard ? 'Да' : 'Нет',
                    'rail_security' => $security === 'no' ? 'Нет' : ($security === '20' ? '20 фут' : '40 фут'),
                    'rail_profit' => $profit,
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
     * Загружаем ЖД маршруты
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
            $cols = ['A','B','C','D','E','F','G','H','I','J','K']; // до K включительно
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
                            'PROPERTY_166' => trim((string)$row['C']),
                            'PROPERTY_168' => trim((string)$row['D']),
                            'PROPERTY_170' => trim((string)$row['E']),
                            'PROPERTY_172' => trim((string)$row['F']),
                            'PROPERTY_174' => trim((string)$row['G']),
                            'PROPERTY_176' => trim((string)$row['H']),
                            'PROPERTY_178' => trim((string)$row['I']),
                            'PROPERTY_180' => trim((string)$row['J']),
                            'PROPERTY_196' => trim((string)$row['K']), // agent
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
     * Загружаем морские маршруты
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
            $cols = ['A','B','C','D','E','F','G','H','I','J']; // до J включительно
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
                        'FIELDS'         => [
                            'NAME'         => trim((string)$row['A']),  // Порт
                            'PROPERTY_126' => trim((string)$row['B']),
                            'PROPERTY_162' => trim((string)$row['C']),
                            'PROPERTY_164' => trim((string)$row['D']),
                            'PROPERTY_132' => trim((string)$row['E']),
                            'PROPERTY_134' => trim((string)$row['F']),
                            'PROPERTY_136' => trim((string)$row['G']),
                            'PROPERTY_138' => trim((string)$row['H']),
                            'PROPERTY_140' => trim((string)$row['I']),
                            'PROPERTY_192' => trim((string)$row['J']), // agent
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
     * Загружаем комбинированные маршруты
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
    
    // Логируем общий результат
    file_put_contents(__DIR__ . '/fetch_transport_data.log', 
        "Total elements fetched: " . count($allElements) . PHP_EOL . 
        "Total transformed: " . count($result) . PHP_EOL, 
        FILE_APPEND | LOCK_EX);
    
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