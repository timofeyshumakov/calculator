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
    public const RAIL_COLUMNS = [
        '20DC (<24t)' => [
            'coc_normal' => 'PROPERTY_166',
            'coc_danger' => 'PROPERTY_168',
            'soc_normal' => 'PROPERTY_166', // SOC использует те же базовые стоимости
            'soc_danger' => 'PROPERTY_168'
        ],
        '20DC (24t-28t)' => [
            'coc_normal' => 'PROPERTY_170',
            'coc_danger' => 'PROPERTY_172',
            'soc_normal' => 'PROPERTY_170',
            'soc_danger' => 'PROPERTY_172'
        ],
        '40HC (28t)' => [
            'coc_normal' => 'PROPERTY_174',
            'coc_danger' => 'PROPERTY_176',
            'soc_normal' => 'PROPERTY_174',
            'soc_danger' => 'PROPERTY_176'
        ]
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
/**
 * Экспорт результатов морских перевозок в Excel
 */
public function exportSeaToExcel()
{
    $input = json_decode(file_get_contents('php://input'), true);
    
    if (empty($input) || !is_array($input)) {
        header('Content-Type: application/json; charset=utf-8');
        echo json_encode(['error' => true, 'message' => 'Нет данных для экспорта']);
        return;
    }
    
    $exportData = $input['export_data'] ?? [];
    $calculationParams = $input['calculation_params'] ?? [];
    $exactMatch = $input['exact_match'] ?? false;
    
    // Если нужно точное соответствие, пересчитываем данные
    if ($exactMatch && !empty($calculationParams)) {
        $_POST = $calculationParams;
        $result = $this->getSeaPerevozki(true);
        if (is_string($result)) {
            $result = json_decode($result, true);
        }
        $exportData = is_array($result) && !isset($result['error']) ? $result : $exportData;
    }
    
    if (empty($exportData)) {
        header('Content-Type: application/json; charset=utf-8');
        echo json_encode(['error' => true, 'message' => 'Нет данных для экспорта']);
        return;
    }
    
    $this->generateSeaExcel($exportData, 'sea_export_' . date('Y-m-d_H-i-s'));
}

/**
 * Экспорт результатов ж/д перевозок в Excel
 */
public function exportRailToExcel()
{
    $input = json_decode(file_get_contents('php://input'), true);
    
    if (empty($input) || !is_array($input)) {
        header('Content-Type: application/json; charset=utf-8');
        echo json_encode(['error' => true, 'message' => 'Нет данных для экспорта']);
        return;
    }
    
    $exportData = $input['export_data'] ?? [];
    $calculationParams = $input['calculation_params'] ?? [];
    $exactMatch = $input['exact_match'] ?? false;
    
    // Если нужно точное соответствие, пересчитываем данные
    if ($exactMatch && !empty($calculationParams)) {
        $_POST = $calculationParams;
        $result = $this->getRailPerevozki(true);
        if (is_string($result)) {
            $result = json_decode($result, true);
        }
        $exportData = is_array($result) && !isset($result['error']) ? $result : $exportData;
    }
    
    if (empty($exportData)) {
        header('Content-Type: application/json; charset=utf-8');
        echo json_encode(['error' => true, 'message' => 'Нет данных для экспорта']);
        return;
    }
    
    $this->generateRailExcel($exportData, 'rail_export_' . date('Y-m-d_H-i-s'));
}

/**
 * Экспорт результатов комбинированных перевозок в Excel
 */
public function exportCombToExcel()
{
    $input = json_decode(file_get_contents('php://input'), true);
    
    if (empty($input) || !is_array($input)) {
        header('Content-Type: application/json; charset=utf-8');
        echo json_encode(['error' => true, 'message' => 'Нет данных для экспорта']);
        return;
    }
    
    $exportData = $input['export_data'] ?? [];
    $calculationParams = $input['calculation_params'] ?? [];
    $exactMatch = $input['exact_match'] ?? false;
    
    // Если нужно точное соответствие, пересчитываем данные
    if ($exactMatch && !empty($calculationParams)) {
        $_POST = $calculationParams;
        $result = $this->getCombPerevozki(true);
        if (is_string($result)) {
            $result = json_decode($result, true);
        }
        $exportData = is_array($result) && !isset($result['error']) ? $result : $exportData;
    }
    
    if (empty($exportData)) {
        header('Content-Type: application/json; charset=utf-8');
        echo json_encode(['error' => true, 'message' => 'Нет данных для экспорта']);
        return;
    }
    
    $this->generateCombExcel($exportData, 'combined_export_' . date('Y-m-d_H-i-s'));
}
    
/**
 * Генерация Excel для морских перевозок с точным соответствием таблице
 */
private function generateSeaExcel($data, $filename)
{
    try {
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        
        // Заголовки точно как в таблице результатов
        $headers = [
            'A' => 'Порт отправления',
            'B' => 'Порт прибытия',
            'C' => 'DROP OFF LOCATION',
            'D' => 'Тип контейнера',
            'E' => 'Собственность контейнера',
            'F' => 'Охрана',
            'G' => 'CAF (%)',
            'H' => 'Стоимость обычного груза, USD',
            'I' => 'Стоимость опасного груза, USD',
            'J' => 'Агент',
            'K' => 'Примечание'
        ];
        
        // Заполняем заголовки
        foreach ($headers as $col => $header) {
            $sheet->setCellValue($col . '1', $header);
            $sheet->getStyle($col . '1')->getFont()->setBold(true);
        }
        
        // Заполняем данные точно как в таблице
        $row = 2;
        foreach ($data as $item) {
            $sheet->setCellValue('A' . $row, $item['sea_pol'] ?? '');
            $sheet->setCellValue('B' . $row, $item['sea_pod'] ?? '');
            $sheet->setCellValue('C' . $row, $item['sea_drop_off_location'] ?? '');
            $sheet->setCellValue('D' . $row, $item['sea_coc'] ?? '');
            $sheet->setCellValue('E' . $row, $item['sea_container_ownership'] ?? '');
            $sheet->setCellValue('F' . $row, $item['sea_security'] ?? 'Нет');
            $sheet->setCellValue('G' . $row, $item['sea_caf_percent'] ?? 0);
            $sheet->setCellValue('H' . $row, $item['cost_total_normal'] ?? 0);
            $sheet->setCellValue('I' . $row, $item['cost_total_danger'] ?? 0);
            $sheet->setCellValue('J' . $row, $item['sea_agent'] ?? '');
            $sheet->setCellValue('K' . $row, $item['sea_remark'] ?? '');
            $row++;
        }
        
        $this->finalizeExcel($spreadsheet, $sheet, $filename);
        
    } catch (\Exception $e) {
        $this->handleExcelError($e, 'морских перевозок');
    }
}

/**
 * Генерация Excel для ж/д перевозок с точным соответствием таблице
 */
private function generateRailExcel($data, $filename)
{
    try {
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        
        // Заголовки точно как в таблице результатов
        $headers = [
            'A' => 'Станция отправления',
            'B' => 'Пункт назначения',
            'C' => 'Тип контейнера',
            'D' => 'Собственность контейнера',
            'E' => 'Охрана',
            'F' => 'Стоимость обычного груза, RUB',
            'G' => 'Стоимость опасного груза, RUB',
            'H' => 'Агент',
            'I' => 'Комментарий'
        ];
        
        // Заполняем заголовки
        foreach ($headers as $col => $header) {
            $sheet->setCellValue($col . '1', $header);
            $sheet->getStyle($col . '1')->getFont()->setBold(true);
        }
        
        // Заполняем данные точно как в таблице
        $row = 2;
        foreach ($data as $item) {
            $sheet->setCellValue('A' . $row, $item['rail_origin'] ?? '');
            $sheet->setCellValue('B' . $row, $item['rail_destination'] ?? '');
            $sheet->setCellValue('C' . $row, $item['rail_coc'] ?? '');
            $sheet->setCellValue('D' . $row, $item['rail_container_ownership'] ?? '');
            $sheet->setCellValue('E' . $row, $item['rail_security'] ?? 'Нет');
            $sheet->setCellValue('F' . $row, $item['cost_total_normal'] ?? 0);
            $sheet->setCellValue('G' . $row, $item['cost_total_danger'] ?? 0);
            $sheet->setCellValue('H' . $row, $item['rail_agent'] ?? '');
            $sheet->setCellValue('I' . $row, $item['rail_remark'] ?? '');
            $row++;
        }
        
        $this->finalizeExcel($spreadsheet, $sheet, $filename);
        
    } catch (\Exception $e) {
        $this->handleExcelError($e, 'ж/д перевозок');
    }
}

/**
 * Генерация Excel для комбинированных перевозок с точным соответствием таблице
 */
private function generateCombExcel($data, $filename)
{
    try {
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        
        // Заголовки точно как в таблице результатов
        $headers = [
            'A' => 'Морской порт отправления',
            'B' => 'Морской порт прибытия',
            'C' => 'DROP OFF LOCATION',
            'D' => 'Тип контейнера',
            'E' => 'Собственность контейнера',
            'F' => 'Охрана',
            'G' => 'Стоимость обычного груза, USD/RUB',
            'H' => 'Стоимость опасного груза, USD/RUB',
            'I' => 'Агент',
            'J' => 'Комментарий'
        ];
        
        // Заполняем заголовки
        foreach ($headers as $col => $header) {
            $sheet->setCellValue($col . '1', $header);
            $sheet->getStyle($col . '1')->getFont()->setBold(true);
        }
        
        // Заполняем данные точно как в таблице
        $row = 2;
        foreach ($data as $item) {
            $sheet->setCellValue('A' . $row, $item['comb_sea_pol'] ?? '');
            $sheet->setCellValue('B' . $row, $item['comb_sea_pod'] ?? '');
            $sheet->setCellValue('C' . $row, $item['comb_drop_off'] ?? '');
            $sheet->setCellValue('D' . $row, $item['comb_coc'] ?? '');
            $sheet->setCellValue('E' . $row, $item['comb_container_ownership'] ?? '');
            $sheet->setCellValue('F' . $row, $item['comb_security'] ?? 'Нет');
            $sheet->setCellValue('G' . $row, $item['cost_total_normal_text'] ?? '');
            $sheet->setCellValue('H' . $row, $item['cost_total_danger_text'] ?? '');
            $sheet->setCellValue('I' . $row, $item['comb_agent'] ?? '');
            $sheet->setCellValue('J' . $row, $item['comb_remark'] ?? '');
            $row++;
        }
        
        $this->finalizeExcel($spreadsheet, $sheet, $filename);
        
    } catch (\Exception $e) {
        $this->handleExcelError($e, 'комбинированных перевозок');
    }
}

/**
 * Финальная обработка Excel файла
 */
private function finalizeExcel($spreadsheet, $sheet, $filename)
{
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
    
    // Форматирование числовых ячеек
    $numericColumns = ['G', 'H', 'I']; // Настройте по мере необходимости
    foreach ($numericColumns as $col) {
        if ($col <= $lastColumn) {
            for ($row = 2; $row <= $lastRow; $row++) {
                $cell = $sheet->getCell($col . $row);
                if (is_numeric($cell->getValue())) {
                    $cell->getStyle()->getNumberFormat()->setFormatCode('#,##0.00');
                }
            }
        }
    }
    
    // Сохраняем файл
    $writer = new Xlsx($spreadsheet);
    
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header('Content-Disposition: attachment;filename="' . $filename . '.xlsx"');
    header('Cache-Control: max-age=0');
    
    $writer->save('php://output');
    exit;
}

/**
 * Обработка ошибок Excel
 */
private function handleExcelError($e, $type)
{
    header('Content-Type: application/json; charset=utf-8');
    echo json_encode([
        'error' => true,
        'message' => 'Ошибка при экспорте ' . $type . ' в Excel: ' . $e->getMessage()
    ]);
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
            // Проверяем, показываем ли оба варианта собственности
            $showBothOwnership = !empty($data[0]['show_both_ownership']) && $data[0]['show_both_ownership'];
            
            if ($showBothOwnership) {
                $headers = [
                    'A' => 'Порт отправления (POL)',
                    'B' => 'Порт прибытия (POD)',
                    'C' => 'DROP OFF LOCATION',
                    'D' => 'Тип контейнера',
                    'E' => 'Собственность контейнера',
                    'F' => 'Опасный груз',
                    'G' => 'Охрана',
                    'H' => 'Стоимость контейнера COC обычный ($)',
                    'I' => 'Стоимость контейнера COC опасный ($)',
                    'J' => 'Стоимость контейнера SOC обычный ($)',
                    'K' => 'Стоимость контейнера SOC опасный ($)',
                    'L' => 'Стоимость DROP OFF ($)',
                    'M' => 'Стоимость охраны ($)',
                    'N' => 'NETTO COC обычный ($)',
                    'O' => 'NETTO COC опасный ($)',
                    'P' => 'NETTO SOC обычный ($)',
                    'Q' => 'NETTO SOC опасный ($)',
                    'R' => 'CAF (%)',
                    'S' => 'Profit ($)',
                    'T' => 'Итоговая стоимость COC обычный ($)',
                    'U' => 'Итоговая стоимость COC опасный ($)',
                    'V' => 'Итоговая стоимость SOC обычный ($)',
                    'W' => 'Итоговая стоимость SOC опасный ($)',
                    'X' => 'Агент',
                    'Y' => 'Примечание',
                ];
            } else {
                $headers = [
                    'A' => 'Порт отправления (POL)',
                    'B' => 'Порт прибытия (POD)',
                    'C' => 'DROP OFF LOCATION',
                    'D' => 'Тип контейнера',
                    'E' => 'Собственность контейнера',
                    'F' => 'Опасный груз',
                    'G' => 'Охрана',
                    'H' => 'Стоимость контейнера обычный ($)',
                    'I' => 'Стоимость контейнера опасный ($)',
                    'J' => 'Стоимость DROP OFF ($)',
                    'K' => 'Стоимость охраны ($)',
                    'L' => 'NETTO обычный ($)',
                    'M' => 'NETTO опасный ($)',
                    'N' => 'CAF (%)',
                    'O' => 'Profit ($)',
                    'P' => 'Итоговая стоимость обычный ($)',
                    'Q' => 'Итоговая стоимость опасный ($)',
                    'R' => 'Агент',
                    'S' => 'Примечание',
                ];
            }
            
            // Заполняем заголовки
            foreach ($headers as $col => $header) {
                $sheet->setCellValue($col . '1', $header);
                $sheet->getStyle($col . '1')->getFont()->setBold(true);
            }
            
            // Заполняем данные
            $row = 2;
            foreach ($data as $item) {
                if ($showBothOwnership) {
                    // Формат с двумя вариантами собственности
                    $sheet->setCellValue('A' . $row, $item['sea_pol'] ?? '');
                    $sheet->setCellValue('B' . $row, $item['sea_pod'] ?? '');
                    $sheet->setCellValue('C' . $row, $item['sea_drop_off_location'] ?? '');
                    $sheet->setCellValue('D' . $row, $item['sea_coc'] ?? '');
                    $sheet->setCellValue('E' . $row, 'COC/SOC'); // Оба варианта
                    $sheet->setCellValue('F' . $row, 'Оба варианта'); // И обычный и опасный
                    $sheet->setCellValue('G' . $row, $item['sea_security'] ?? 'Нет');
                    $sheet->setCellValue('H' . $row, $item['coc_container_cost_normal'] ?? 0);
                    $sheet->setCellValue('I' . $row, $item['coc_container_cost_danger'] ?? 0);
                    $sheet->setCellValue('J' . $row, $item['soc_container_cost_normal'] ?? 0);
                    $sheet->setCellValue('K' . $row, $item['soc_container_cost_danger'] ?? 0);
                    $sheet->setCellValue('L' . $row, $item['cost_drop_off'] ?? 0);
                    $sheet->setCellValue('M' . $row, $item['cost_security'] ?? 0);
                    $sheet->setCellValue('N' . $row, $item['coc_netto_normal'] ?? 0);
                    $sheet->setCellValue('O' . $row, $item['coc_netto_danger'] ?? 0);
                    $sheet->setCellValue('P' . $row, $item['soc_netto_normal'] ?? 0);
                    $sheet->setCellValue('Q' . $row, $item['soc_netto_danger'] ?? 0);
                    $sheet->setCellValue('R' . $row, $item['sea_caf_percent'] ?? 0);
                    $sheet->setCellValue('S' . $row, $item['sea_profit'] ?? 0);
                    $sheet->setCellValue('T' . $row, $item['coc_total_normal'] ?? 0);
                    $sheet->setCellValue('U' . $row, $item['coc_total_danger'] ?? 0);
                    $sheet->setCellValue('V' . $row, $item['soc_total_normal'] ?? 0);
                    $sheet->setCellValue('W' . $row, $item['soc_total_danger'] ?? 0);
                    $sheet->setCellValue('X' . $row, $item['sea_agent'] ?? '');
                    $sheet->setCellValue('Y' . $row, $item['sea_remark'] ?? '');
                } else {
                    // Формат с одним вариантом собственности
                    $sheet->setCellValue('A' . $row, $item['sea_pol'] ?? '');
                    $sheet->setCellValue('B' . $row, $item['sea_pod'] ?? '');
                    $sheet->setCellValue('C' . $row, $item['sea_drop_off_location'] ?? '');
                    $sheet->setCellValue('D' . $row, $item['sea_coc'] ?? '');
                    $sheet->setCellValue('E' . $row, $item['sea_container_ownership'] ?? '');
                    $sheet->setCellValue('F' . $row, $item['sea_hazard'] ?? 'Нет');
                    $sheet->setCellValue('G' . $row, $item['sea_security'] ?? 'Нет');
                    $sheet->setCellValue('H' . $row, $item['cost_container_normal'] ?? 0);
                    $sheet->setCellValue('I' . $row, $item['cost_container_danger'] ?? 0);
                    $sheet->setCellValue('J' . $row, $item['cost_drop_off'] ?? 0);
                    $sheet->setCellValue('K' . $row, $item['cost_security'] ?? 0);
                    $sheet->setCellValue('L' . $row, $item['cost_netto_normal'] ?? 0);
                    $sheet->setCellValue('M' . $row, $item['cost_netto_danger'] ?? 0);
                    $sheet->setCellValue('N' . $row, $item['sea_caf_percent'] ?? 0);
                    $sheet->setCellValue('O' . $row, $item['sea_profit'] ?? 0);
                    $sheet->setCellValue('P' . $row, $item['cost_total_normal'] ?? 0);
                    $sheet->setCellValue('Q' . $row, $item['cost_total_danger'] ?? 0);
                    $sheet->setCellValue('R' . $row, $item['sea_agent'] ?? '');
                    $sheet->setCellValue('S' . $row, $item['sea_remark'] ?? '');
                }
                $row++;
            }
            
        } elseif ($type === 'rail') {
$showBothOwnership = !empty($data[0]['show_both_ownership']) && $data[0]['show_both_ownership'];
            $showBothHazard = !empty($data[0]['show_both_hazard_in_columns']) && $data[0]['show_both_hazard_in_columns'];
            
            if ($showBothOwnership && $showBothHazard) {
                // Формат с двумя вариантами собственности и двумя типами опасности
                $headers = [
                    'A' => 'Станция отправления',
                    'B' => 'Станция прибытия',
                    'C' => 'Тип контейнера',
                    'D' => 'Собственность контейнера',
                    'E' => 'Опасный груз',
                    'F' => 'Охрана',
                    'G' => 'Прибыль (₽)',
                    'H' => 'Стоимость COC 20DC <24t обычный (₽)',
                    'I' => 'Стоимость COC 20DC <24t опасный (₽)',
                    'J' => 'Стоимость COC 20DC 24t-28t обычный (₽)',
                    'K' => 'Стоимость COC 20DC 24t-28t опасный (₽)',
                    'L' => 'Стоимость COC 40HC 28t обычный (₽)',
                    'M' => 'Стоимость COC 40HC 28t опасный (₽)',
                    'N' => 'Стоимость SOC 20DC <24t обычный (₽)',
                    'O' => 'Стоимость SOC 20DC <24t опасный (₽)',
                    'P' => 'Стоимость SOC 20DC 24t-28t обычный (₽)',
                    'Q' => 'Стоимость SOC 20DC 24t-28t опасный (₽)',
                    'R' => 'Стоимость SOC 40HC 28t обычный (₽)',
                    'S' => 'Стоимость SOC 40HC 28t опасный (₽)',
                    'T' => 'Стоимость охраны (₽)',
                    'U' => 'Итог COC 20DC <24t обычный (₽)',
                    'V' => 'Итог COC 20DC <24t опасный (₽)',
                    'W' => 'Итог COC 20DC 24t-28t обычный (₽)',
                    'X' => 'Итог COC 20DC 24t-28t опасный (₽)',
                    'Y' => 'Итог COC 40HC 28t обычный (₽)',
                    'Z' => 'Итог COC 40HC 28t опасный (₽)',
                    'AA' => 'Итог SOC 20DC <24t обычный (₽)',
                    'AB' => 'Итог SOC 20DC <24t опасный (₽)',
                    'AC' => 'Итог SOC 20DC 24t-28t обычный (₽)',
                    'AD' => 'Итог SOC 20DC 24t-28t опасный (₽)',
                    'AE' => 'Итог SOC 40HC 28t обычный (₽)',
                    'AF' => 'Итог SOC 40HC 28t опасный (₽)',
                    'AG' => 'Агент',
                ];
            } elseif ($showBothHazard) {
                // Формат с одним вариантом собственности, но двумя типами опасности
                $headers = [
                    'A' => 'Станция отправления',
                    'B' => 'Станция прибытия',
                    'C' => 'Тип контейнера',
                    'D' => 'Собственность контейнера',
                    'E' => 'Опасный груз',
                    'F' => 'Охрана',
                    'G' => 'Прибыль (₽)',
                    'H' => 'Стоимость 20DC <24t обычный (₽)',
                    'I' => 'Стоимость 20DC <24t опасный (₽)',
                    'J' => 'Стоимость 20DC 24t-28t обычный (₽)',
                    'K' => 'Стоимость 20DC 24t-28t опасный (₽)',
                    'L' => 'Стоимость 40HC 28t обычный (₽)',
                    'M' => 'Стоимость 40HC 28т опасный (₽)',
                    'N' => 'Стоимость охраны (₽)',
                    'O' => 'Итог 20DC <24t обычный (₽)',
                    'P' => 'Итог 20DC <24t опасный (₽)',
                    'Q' => 'Итог 20DC 24t-28t обычный (₽)',
                    'R' => 'Итог 20DC 24t-28t опасный (₽)',
                    'S' => 'Итог 40HC 28t обычный (₽)',
                    'T' => 'Итог 40HC 28t опасный (₽)',
                    'U' => 'Агент',
                ];
            } else {
                // Стандартный формат (только один тип опасности)
                $headers = [
                    'A' => 'Станция отправления',
                    'B' => 'Станция прибытия',
                    'C' => 'Тип контейнера',
                    'D' => 'Собственность контейнера',
                    'E' => 'Опасный груз',
                    'F' => 'Охрана',
                    'G' => 'Прибыль (₽)',
                    'H' => 'Стоимость 20DC <24t (₽)',
                    'I' => 'Стоимость 20DC 24t-28t (₽)',
                    'J' => 'Стоимость 40HC 28t (₽)',
                    'K' => 'Стоимость охраны (₽)',
                    'L' => 'Итог 20DC <24t (₽)',
                    'M' => 'Итог 20DC 24t-28t (₽)',
                    'N' => 'Итог 40HC 28t (₽)',
                    'O' => 'Агент',
                ];
            }
            
            // Заполняем заголовки
            foreach ($headers as $col => $header) {
                $sheet->setCellValue($col . '1', $header);
                $sheet->getStyle($col . '1')->getFont()->setBold(true);
            }
            
            // Заполняем данные
            $row = 2;
            foreach ($data as $item) {
                if ($showBothOwnership && $showBothHazard) {
                    // Формат с двумя вариантами собственности и двумя типами опасности
                    $sheet->setCellValue('A' . $row, $item['rail_origin'] ?? '');
                    $sheet->setCellValue('B' . $row, $item['rail_destination'] ?? '');
                    $sheet->setCellValue('C' . $row, $item['rail_coc'] ?? '');
                    $sheet->setCellValue('D' . $row, 'COC/SOC');
                    $sheet->setCellValue('E' . $row, 'Оба варианта');
                    $sheet->setCellValue('F' . $row, $item['rail_security'] ?? 'Нет');
                    $sheet->setCellValue('G' . $row, $item['rail_profit'] ?? 0);
                    
                    // COC обычный
                    $sheet->setCellValue('H' . $row, $item['coc_cost_base_20'] ?? 0);
                    $sheet->setCellValue('I' . $row, $item['coc_cost_base_20_danger'] ?? 0);
                    $sheet->setCellValue('J' . $row, $item['coc_cost_base_20_28'] ?? 0);
                    $sheet->setCellValue('K' . $row, $item['coc_cost_base_20_28_danger'] ?? 0);
                    $sheet->setCellValue('L' . $row, $item['coc_cost_base_40'] ?? 0);
                    $sheet->setCellValue('M' . $row, $item['coc_cost_base_40_danger'] ?? 0);
                    
                    // SOC обычный
                    $sheet->setCellValue('N' . $row, $item['soc_cost_base_20'] ?? 0);
                    $sheet->setCellValue('O' . $row, $item['soc_cost_base_20_danger'] ?? 0);
                    $sheet->setCellValue('P' . $row, $item['soc_cost_base_20_28'] ?? 0);
                    $sheet->setCellValue('Q' . $row, $item['soc_cost_base_20_28_danger'] ?? 0);
                    $sheet->setCellValue('R' . $row, $item['soc_cost_base_40'] ?? 0);
                    $sheet->setCellValue('S' . $row, $item['soc_cost_base_40_danger'] ?? 0);
                    
                    $sheet->setCellValue('T' . $row, $item['cost_security'] ?? 0);
                    
                    // COC итог
                    $sheet->setCellValue('U' . $row, $item['coc_cost_total_20'] ?? 0);
                    $sheet->setCellValue('V' . $row, $item['coc_cost_total_20_danger'] ?? 0);
                    $sheet->setCellValue('W' . $row, $item['coc_cost_total_20_28'] ?? 0);
                    $sheet->setCellValue('X' . $row, $item['coc_cost_total_20_28_danger'] ?? 0);
                    $sheet->setCellValue('Y' . $row, $item['coc_cost_total_40'] ?? 0);
                    $sheet->setCellValue('Z' . $row, $item['coc_cost_total_40_danger'] ?? 0);
                    
                    // SOC итог
                    $sheet->setCellValue('AA' . $row, $item['soc_cost_total_20'] ?? 0);
                    $sheet->setCellValue('AB' . $row, $item['soc_cost_total_20_danger'] ?? 0);
                    $sheet->setCellValue('AC' . $row, $item['soc_cost_total_20_28'] ?? 0);
                    $sheet->setCellValue('AD' . $row, $item['soc_cost_total_20_28_danger'] ?? 0);
                    $sheet->setCellValue('AE' . $row, $item['soc_cost_total_40'] ?? 0);
                    $sheet->setCellValue('AF' . $row, $item['soc_cost_total_40_danger'] ?? 0);
                    
                    $sheet->setCellValue('AG' . $row, $item['rail_agent'] ?? '');
                    
                } elseif ($showBothHazard) {
                    // Формат с одним вариантом собственности, но двумя типами опасности
                    $sheet->setCellValue('A' . $row, $item['rail_origin'] ?? '');
                    $sheet->setCellValue('B' . $row, $item['rail_destination'] ?? '');
                    $sheet->setCellValue('C' . $row, $item['rail_coc'] ?? '');
                    $sheet->setCellValue('D' . $row, $item['rail_container_ownership'] ?? '');
                    $sheet->setCellValue('E' . $row, 'Оба варианта');
                    $sheet->setCellValue('F' . $row, $item['rail_security'] ?? 'Нет');
                    $sheet->setCellValue('G' . $row, $item['rail_profit'] ?? 0);
                    
                    // Базовые стоимости
                    $sheet->setCellValue('H' . $row, $item['cost_base_20'] ?? 0);
                    $sheet->setCellValue('I' . $row, $item['cost_base_20_danger'] ?? 0);
                    $sheet->setCellValue('J' . $row, $item['cost_base_20_28'] ?? 0);
                    $sheet->setCellValue('K' . $row, $item['cost_base_20_28_danger'] ?? 0);
                    $sheet->setCellValue('L' . $row, $item['cost_base_40'] ?? 0);
                    $sheet->setCellValue('M' . $row, $item['cost_base_40_danger'] ?? 0);
                    
                    $sheet->setCellValue('N' . $row, $item['cost_security'] ?? 0);
                    
                    // Итоговые стоимости
                    $sheet->setCellValue('O' . $row, $item['cost_total_20'] ?? 0);
                    $sheet->setCellValue('P' . $row, $item['cost_total_20_danger'] ?? 0);
                    $sheet->setCellValue('Q' . $row, $item['cost_total_20_28'] ?? 0);
                    $sheet->setCellValue('R' . $row, $item['cost_total_20_28_danger'] ?? 0);
                    $sheet->setCellValue('S' . $row, $item['cost_total_40'] ?? 0);
                    $sheet->setCellValue('T' . $row, $item['cost_total_40_danger'] ?? 0);
                    
                    $sheet->setCellValue('U' . $row, $item['rail_agent'] ?? '');
                    
                } else {
                    // Стандартный формат
                    $sheet->setCellValue('A' . $row, $item['rail_origin'] ?? '');
                    $sheet->setCellValue('B' . $row, $item['rail_destination'] ?? '');
                    $sheet->setCellValue('C' . $row, $item['rail_coc'] ?? '');
                    $sheet->setCellValue('D' . $row, $item['rail_container_ownership'] ?? '');
                    $sheet->setCellValue('E' . $row, $item['rail_hazard'] ?? 'Нет');
                    $sheet->setCellValue('F' . $row, $item['rail_security'] ?? 'Нет');
                    $sheet->setCellValue('G' . $row, $item['rail_profit'] ?? 0);
                    
                    $sheet->setCellValue('H' . $row, $item['cost_base_20'] ?? 0);
                    $sheet->setCellValue('I' . $row, $item['cost_base_20_28'] ?? 0);
                    $sheet->setCellValue('J' . $row, $item['cost_base_40'] ?? 0);
                    $sheet->setCellValue('K' . $row, $item['cost_security'] ?? 0);
                    
                    $sheet->setCellValue('L' . $row, $item['cost_total_20'] ?? 0);
                    $sheet->setCellValue('M' . $row, $item['cost_total_20_28'] ?? 0);
                    $sheet->setCellValue('N' . $row, $item['cost_total_40'] ?? 0);
                    
                    $sheet->setCellValue('O' . $row, $item['rail_agent'] ?? '');
                }
                $row++;
            }
            
        } elseif ($type === 'combined') {
            // Аналогичные изменения для комбинированных перевозок...
            // (нужно будет обновить и этот блок по аналогии)
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
    /**
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
            28, 
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
                28, 
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
        $cafPercent = floatval($params['sea_caf'] ?? 0);
        $profit = floatval($params['sea_profit'] ?? 0);
        $containerOwnership = $params['sea_container_ownership'] ?? 'no';
        $containerType = $params['sea_coc'] ?? '';
        
        // Определяем типы контейнеров для отображения
        $containerTypesToShow = empty($containerType) ? ['20DC', '40HC'] : [$containerType];
        
        // Флаг для раздельного отображения COC/SOC
        $showBothOwnership = ($containerOwnership === 'no');
        
        // Обрабатываем каждую найденную запись
        foreach ($seaPerevozki as $value) {
            foreach ($containerTypesToShow as $cType) {
                $is20GP = ($cType === '20DC');
    //echo json_encode($cType, JSON_UNESCAPED_UNICODE);
    //return json_encode($cType, JSON_UNESCAPED_UNICODE);
                // Получаем стоимости
                $costs = $this->getSeaCosts($value, $is20GP);
                
                // Проверяем наличие цен для выбранных параметров
                $shouldSkip = false;
                
                if ($showBothOwnership) {
                    // Проверяем оба варианта: COC и SOC
                    if ($isHazard) {
                        // Для опасного груза проверяем цены COC
                        if (empty($costs['coc_danger']) || floatval($costs['coc_danger']) <= 0) {
                            $shouldSkip = true;
                        }
                    } else {
                        // Для обычного груза проверяем цены COC и SOC
                        if ((empty($costs['coc_normal']) || floatval($costs['coc_normal']) <= 0) &&
                            (empty($costs['soc_normal']) || floatval($costs['soc_normal']) <= 0)) {
                            $shouldSkip = true;
                        }
                    }
                } else {
                    // Проверяем один вариант собственности
                    if ($containerOwnership === 'coc') {
                        if ($isHazard) {
                            // COC опасный
                            if (empty($costs['coc_danger']) || floatval($costs['coc_danger']) <= 0) {
                                $shouldSkip = true;
                            }
                        } else {
                            // COC обычный
                            if (empty($costs['coc_normal']) || floatval($costs['coc_normal']) <= 0) {
                                $shouldSkip = true;
                            }
                        }
                    } elseif ($containerOwnership === 'soc') {
                        if ($isHazard) {
                            // SOC опасный (используем COC цены)
                            if (empty($costs['coc_danger']) || floatval($costs['coc_danger']) <= 0) {
                                $shouldSkip = true;
                            }
                        } else {
                            // SOC обычный
                            if (empty($costs['soc_normal']) || floatval($costs['soc_normal']) <= 0) {
                                $shouldSkip = true;
                            }
                        }
                    }
                }
                
                // Пропускаем строку если нет цены
                if ($shouldSkip) {
                    continue;
                }
                
                $dropOffCost = $is20GP ? $costs['drop_off_20'] : $costs['drop_off_40'];

                // Если показываем оба варианта собственности
                if ($showBothOwnership) {
                    // Проверяем наличие цены для COC
                    $hasCOCCost = !empty($costs['coc_normal']) && floatval($costs['coc_normal']) > 0;
                    $hasSOCSCost = !empty($costs['soc_normal']) && floatval($costs['soc_normal']) > 0;
                    
                    // Добавляем COC вариант только если есть цена
                    if ($hasCOCCost) {
                        $result[] = $this->createSeaResultItem(
                            $value, $cType, 'COC', 'Оба варианта', 
                            $cafPercent, $profit, $dropOffCost,
                            $costs['coc_normal'], $costs['coc_danger']
                        );
                    }
                    
                    // Добавляем SOC вариант только если есть цена
                    if ($hasSOCSCost) {
                        $result[] = $this->createSeaResultItem(
                            $value, $cType, 'SOC', 'Оба варианта',
                            $cafPercent, $profit, $dropOffCost,
                            $costs['soc_normal'], $costs['coc_danger'], // Для опасного используем COC стоимость
                            $costs['soc_normal']
                        );
                    }
                } else {
                    // Один вариант собственности
                    $selectedType = $containerOwnership === 'coc' ? 'COC' : 'SOC';
                    
                    // Получаем стоимость контейнера
                    $containerCost = 0;
                    if ($containerOwnership === 'coc') {
                        $containerCost = $is20GP ? $costs['coc_normal'] : $costs['coc_40_normal'];
                    } else {
                        $containerCost = $is20GP ? $costs['soc_normal'] : $costs['soc_40_normal'];
                    }
                    
                    // Проверяем что стоимость не нулевая
                    if (empty($containerCost) || floatval($containerCost) <= 0) {
                        continue;
                    }
                    
                    if (!$isHazard) {
                        // Обычный груз - показываем оба варианта опасности
                        // Проверяем наличие цены для опасного груза
                        $dangerCost = $is20GP ? $costs['coc_danger'] : $costs['coc_40_danger'];
                        
                        if (empty($dangerCost) || floatval($dangerCost) <= 0) {
                            // Если нет цены для опасного груза, показываем только обычный
                            $result[] = $this->createSeaResultItem(
                                $value, $cType, $selectedType, 'Нет',
                                $cafPercent, $profit, $dropOffCost,
                                $containerCost, 0
                            );
                        } else {
                            $result[] = $this->createSeaResultItem(
                                $value, $cType, $selectedType, 'Нет',
                                $cafPercent, $profit, $dropOffCost,
                                $containerCost, $dangerCost
                            );
                        }
                    } else {
                        // Опасный груз - показываем только опасный
                        $dangerCost = $is20GP ? $costs['coc_danger'] : $costs['coc_40_danger'];
                        
                        // Проверяем наличие цены для опасного груза
                        if (empty($dangerCost) || floatval($dangerCost) <= 0) {
                            continue; // Пропускаем если нет цены для опасного груза
                        }
                        
                        $nettoDanger = ceil($dangerCost);
                        $totalDanger = ceil($dangerCost * (1 + $cafPercent / 100) + $profit);
                        
                        $result[] = [
                            'sea_pol' => $value['POL'] ?? '',
                            'sea_pod' => $value['POD'] ?? '',
                            'sea_drop_off_location' => $value['DROP_OFF_LOCATION'] ?? '',
                            'sea_coc' => $cType,
                            'sea_container_ownership' => $selectedType,
                            'sea_agent' => $value['AGENT'] ?? '',
                            'sea_remark' => $value['REMARK'] ?? '',
                            'sea_hazard' => 'Да',
                            'sea_caf_percent' => $cafPercent,
                            'sea_profit' => $profit,
                            'cost_container' => $dangerCost,
                            'cost_netto' => $nettoDanger,
                            'cost_total' => $totalDanger,
                            'is20GP' => $is20GP,
                            'cost_drop_off' => $dropOffCost,
                            'show_both_ownership' => false,
                            'show_both_hazard_in_columns' => false
                        ];
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
    
    return [
        'coc_normal' => ceil(floatval($is20GP ? ($data['COC_20GP'] ?? 0) : ($data['COC_40HC'] ?? 0))),
        'coc_danger' => ceil(floatval($is20GP ? ($data['OPASNYY_20DC'] ?? 0) : ($data['OPASNYY_40HC'] ?? 0))),
        'soc_normal' => ceil(floatval($is20GP ? ($data['SOC_20GP'] ?? 0) : ($data['SOC_40HC'] ?? 0))),
        'coc_40_normal' => ceil(floatval($data['COC_40HC'] ?? 0)),
        'coc_40_danger' => ceil(floatval($data['OPASNYY_40HC'] ?? 0)),
        'soc_40_normal' => ceil(floatval($data['SOC_40HC'] ?? 0)),
        'drop_off_20' => ceil(floatval($data['DROP_OFF_20GP'] ?? 0)),
        'drop_off_40' => ceil(floatval($data['DROP_OFF_40HC'] ?? 0))
    ];
}
/**
 * Создает элемент результата для морских перевозок
 */
private function createSeaResultItem($data, $containerType, $ownership, $hazardType, 
                                     $cafPercent, $profit, $dropOffCost, 
                                     $normalCost, $dangerCost, $socNormalCost = null): array {
    
    $isSOC = ($ownership === 'SOC');
    $containerCost = $isSOC && $socNormalCost ? $socNormalCost : $normalCost;
    
    // Проверяем что базовая стоимость не нулевая
    if (empty($containerCost) || floatval($containerCost) <= 0) {
        return []; // Возвращаем пустой массив
    }
    
    // Расчет для обычного груза
    $nettoNormal = ceil($containerCost + $dropOffCost);
    $totalNormal = ceil($nettoNormal * (1 + $cafPercent / 100) + $profit);
    
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
        'cost_container_normal' => $containerCost,
        'cost_netto_normal' => $nettoNormal,
        'cost_total_normal' => $totalNormal,
        'cost_drop_off' => $dropOffCost,
        'show_both_ownership' => false,
        'show_both_hazard_in_columns' => ($dangerCost > 0) // Показывать оба варианта опасности только если есть цена для опасного груза
    ];
    
    // Добавляем расчет для опасного груза если цена указана
    if (!empty($dangerCost) && floatval($dangerCost) > 0) {
        $nettoDanger = ceil($dangerCost + $dropOffCost);
        $totalDanger = ceil($nettoDanger * (1 + $cafPercent / 100) + $profit);
        
        $resultItem['cost_container_danger'] = $dangerCost;
        $resultItem['cost_netto_danger'] = $nettoDanger;
        $resultItem['cost_total_danger'] = $totalDanger;
    }
    
    return $resultItem;
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
            // Получаем стоимости для выбранного типа контейнера
            $normalCost = $this->getRailCostForContainerType($cocType, $value, false);
            $dangerCost = $this->getRailCostForContainerType($cocType, $value, true);
            
            // Получаем стоимость охраны для выбранного типа контейнера
            $securityCost = $this->getSecurityCostForContainerType($value, $security, $cocType);
            
            // Если собственность контейнера не выбрана ('no') - показываем оба варианта в ОТДЕЛЬНЫХ РЯДАХ
            if ($containerOwnership === 'no') {
                // Ряд для COC
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
                    'cost_base_normal' => $normalCost,
                    'cost_total_normal' => ceil($normalCost + $securityCost + $profit),
                    
                    // Опасный груз
                    'cost_base_danger' => $dangerCost,
                    'cost_total_danger' => ceil($dangerCost + $securityCost + $profit),
                    
                    // Общие поля
                    'cost_security' => $securityCost,
                    'show_both_ownership' => false, // Теперь false, так как показываем в отдельных рядах
                    'show_both_hazard_in_columns' => true
                ];

                // Ряд для SOC
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
                    'cost_base_normal' => $normalCost,
                    'cost_total_normal' => ceil($normalCost + $securityCost + $profit),
                    
                    // Опасный груз
                    'cost_base_danger' => $dangerCost,
                    'cost_total_danger' => ceil($dangerCost + $securityCost + $profit),
                    
                    // Общие поля
                    'cost_security' => $securityCost,
                    'show_both_ownership' => false, // Теперь false, так как показываем в отдельных рядах
                    'show_both_hazard_in_columns' => true
                ];

                // Добавляем оба ряда в результат
                $result[] = $resultItemCOC;
                $result[] = $resultItemSOC;
                
            } else {
                // Если выбрана конкретная собственность контейнера
                $displayContainerType = $containerOwnership === 'coc' ? 'COC' : 'SOC';
                
                if (!$isHazard) {
                    // Пользователь выбрал обычный груз - показываем оба варианта (обычный и опасный)
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
                        'cost_base_normal' => $normalCost,
                        'cost_total_normal' => ceil($normalCost + $securityCost + $profit),
                        
                        // Опасный груз
                        'cost_base_danger' => $dangerCost,
                        'cost_total_danger' => ceil($dangerCost + $securityCost + $profit),
                        
                        // Общие поля
                        'cost_security' => $securityCost,
                        'show_both_ownership' => false,
                        'show_both_hazard_in_columns' => true
                    ];

                    $result[] = $resultItem;
                    
                } else {
                    // Пользователь выбрал опасный груз - показываем только опасный
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
     * Получает стоимость для конкретного типа контейнера
     */
    private function getRailCostForContainerType($containerType, $data, $isDanger = false): float {
        if ($isDanger) {
            // Стоимость для опасного груза
            switch ($containerType) {
                case '20DC (<24t)':
                case '20DC':
                    return ceil(floatval($data['OPASNYY_20DC_24'] ?? 0));
                case '20DC (24t-28t)':
                    return ceil(floatval($data['OPASNYY_DC20_24T_28T'] ?? 0));
                case '40HC (28t)':
                case '40HC':
                    return ceil(floatval($data['OPASNYY_HC40_28T'] ?? 0));
                default:
                    return 0;
            }
        } else {
            // Базовая стоимость для обычного груза
            switch ($containerType) {
                case '20DC (<24t)':
                case '20DC':
                    return ceil(floatval($data['DC20_24'] ?? 0));
                case '20DC (24t-28t)':
                    return ceil(floatval($data['DC20_24T_28T'] ?? 0));
                case '40HC (28t)':
                case '40HC':
                    return ceil(floatval($data['HC40_28T'] ?? 0));
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
        $is40HC = ($containerType === '40HC (28t)');
        
        if (($security === '20' && !$is40HC) || ($security === '40' && $is40HC)) {
            $securityField = $is40HC ? 'OKHRANA_40_FUT' : 'OKHRANA_20_FUT';
            return ceil(floatval($data[$securityField] ?? 0));
        }
        
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
 * Рассчитывает итоговые стоимости
 */
private function calculateTotalCosts($baseCosts, $securityCost, $profit): array {
    return [
        '20' => ceil($baseCosts['20'] + $securityCost + $profit),
        '20_28' => ceil($baseCosts['20_28'] + $securityCost + $profit),
        '40' => ceil($baseCosts['40'] + $securityCost + $profit)
    ];
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
 * Получаем комбинированные маршруты
 *
 * @return [type]
 */
public function getCombPerevozki() {
    header('Content-Type: application/json; charset=utf-8');
    $result = [];
    $params = $_POST;
    
    try {
        // Получаем порт отправления (порт POL из морских перевозок)
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
            28, 
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
            32,
            self::COMB_TRANSPORT_MAP,
            []
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

        // Для каждой найденной морской перевозки
        foreach ($seaPerevozki as $seaValue) {
            // Получаем морской порт прибытия (POD)
            $seaPod = $seaValue['POD'] ?? '';
            
            // Фильтруем комбинированные перевозки по морскому порту прибытия
            $filteredBySeaPod = [];
            foreach ($combPerevozki as $combItem) {
                // Проверяем, что POL в комбинированных соответствует POD морских
                if (($combItem['POL'] ?? '') === $seaPod) {
                    $filteredBySeaPod[] = $combItem;
                }
            }
            
            if (empty($filteredBySeaPod)) {
                continue; // Нет комбинированных перевозок для этого морского POD
            }
            
            // Для каждой подходящей комбинированной перевозки
            foreach ($filteredBySeaPod as $combValue) {
                $railStartStation = $combValue['STANTSIYA_OTPRAVLENIYA'] ?? '';
                $railDestStation = $combValue['STANTSIYA_NAZNACHENIYA'] ?? '';
                $combDestPoint = $combValue['PUNKT_NAZNACHENIYA'] ?? '';
                
                if (empty($railStartStation) || empty($railDestStation)) {
                    continue;
                }
                
                // Получаем ж/д перевозки для станции отправления
                $railData = self::fetchTransportData(
                    30, 
                    self::ZHD_TRANSPORT_MAP,
                    [
                        '=NAME' => $railStartStation,
                    ]
                );
                
                // Фильтруем по станции назначения
                $filteredRailData = [];
                foreach ($railData as $railItem) {
                    if (($railItem['POD'] ?? '') === $railDestStation) {
                        $filteredRailData[] = $railItem;
                    }
                }
                
                if (empty($filteredRailData)) {
                    continue; // Нет подходящих ж/д маршрутов
                }
                
                // Для каждой подходящей ж/д перевозки
                foreach ($filteredRailData as $railValue) {

                    // Если собственность контейнера не выбрана ('no') - показываем оба варианта в ОТДЕЛЬНЫХ РЯДАХ
                    if ($containerOwnership === 'no') {
                        // ===== РЯД ДЛЯ COC =====
                        
                        // Определяем тип контейнера для морской части
                        $is40HC = ($cocType === '40HC (28t)');
                        
                        // Получаем стоимости для COC
                        $cocCost = $is40HC ? ceil(floatval($seaValue['COC_40HC'] ?? 0)) : ceil(floatval($seaValue['COC_20GP'] ?? 0));
                        $cocCostDanger = $is40HC ? ceil(floatval($seaValue['OPASNYY_40HC'] ?? 0)) : ceil(floatval($seaValue['OPASNYY_20DC'] ?? 0));
                        
                        // DROP OFF стоимости
                        $dropOffCost = $is40HC ? ceil(floatval($seaValue['DROP_OFF_40HC'] ?? 0)) : ceil(floatval($seaValue['DROP_OFF_20GP'] ?? 0));
                        
                        // Стоимость охраны для морской части
                        $securityCostSea = 0;
                        if ($security === '20' && !$is40HC) {
                            $securityCostSea = ceil(floatval($seaValue['OKHRANA_20_FUT'] ?? 0));
                        } elseif ($security === '40' && $is40HC) {
                            $securityCostSea = ceil(floatval($seaValue['OKHRANA_40_FUT'] ?? 0));
                        }
                        
                        // CAF процент
                        $cafPercent = floatval($seaValue['CAF_KONVERT'] ?? 0);
                        
                        // Расчет для обычного груза COC
                        $cocNetto = ceil($dropOffCost + $cocCost);
                        $costSeaCocNormal = ceil(($dropOffCost + $cocNetto) * (1 + $cafPercent / 100) + $securityCostSea + $seaProfit);

                        // Расчет для опасного груза COC
                        $cocNettoDanger = ceil($dropOffCost + $cocCostDanger);
                        $costSeaCocDanger = ceil(($dropOffCost + $cocNettoDanger) * (1 + $cafPercent / 100) + $securityCostSea + $seaProfit);
                        
                        // ===== Ж/Д ЧАСТЬ ДЛЯ COC =====
                        $railCostNormal = $this->getRailCostForContainerType($cocType, $railValue, false);
                        $railCostDanger = $this->getRailCostForContainerType($cocType, $railValue, true);
                        
                        // Стоимость охраны для ЖД части
                        $securityCostRail = $this->getSecurityCostForContainerType($railValue, $security, $cocType);

                        $costRailCocNormal = ceil($railCostNormal + $securityCostRail + $railProfit);
                        $costRailCocDanger = ceil($railCostDanger + $securityCostRail + $railProfit);
                        
                        // Общие стоимости для COC
                        $totalCocNormal = $costSeaCocNormal + $costRailCocNormal;
                        $totalCocDanger = $costSeaCocDanger + $costRailCocDanger;
                        
                        $resultItemCOC = [
                            'comb_sea_pol' => $seaValue['POL'] ?? '',
                            'comb_sea_pod' => $seaValue['POD'] ?? '',
                            'comb_rail_start' => $railStartStation,
                            'comb_rail_dest' => $railDestStation,
                            'comb_drop_off' => $seaValue['DROP_OFF_LOCATION'] ?? '',
                            'comb_transshipment_port' => $combValue['PUNKT_OTPRAVLENIYA'] ?? '',
                            'comb_coc' => $cocType,
                            'comb_container_ownership' => 'COC',
                            'comb_hazard' => 'Оба варианта',
                            'comb_security' => $security === 'no' ? 'Нет' : ($security === '20' ? '20 фут' : '40 фут'),
                            'comb_agent' => trim(($seaValue['AGENT'] ?? '') . '; ' . ($railValue['AGENT'] ?? '')),
                            'comb_remark' => $this->getCombinedRemark($seaValue, $combValue, $railValue),
                            
                            // Морская часть - обычный груз
                            'cost_sea_normal' => $costSeaCocNormal,
                            'container_cost_normal' => $cocCost,
                            'drop_off_cost_normal' => $dropOffCost,
                            'caf_percent_normal' => $cafPercent,
                            'security_cost_sea_normal' => $securityCostSea,
                            
                            // Морская часть - опасный груз
                            'cost_sea_danger' => $costSeaCocDanger,
                            'container_cost_danger' => $cocCostDanger,
                            'drop_off_cost_danger' => $dropOffCost,
                            'caf_percent_danger' => $cafPercent,
                            'security_cost_sea_danger' => $securityCostSea,
                            
                            // Ж/Д часть - обычный груз
                            'cost_rail_normal' => $costRailCocNormal,
                            'rail_base_cost_normal' => $railCostNormal,
                            'security_cost_rail_normal' => $securityCostRail,
                            
                            // Ж/Д часть - опасный груз
                            'cost_rail_danger' => $costRailCocDanger,
                            'rail_base_cost_danger' => $railCostDanger,
                            'security_cost_rail_danger' => $securityCostRail,

                            'cost_total_normal_text' => $costSeaCocNormal . '$ + ' . $costRailCocNormal . ' руб',
                            'cost_total_danger_text' => $costSeaCocDanger . '$ + ' . $costRailCocDanger . ' руб',
                            
                            'show_both_ownership' => false,
                            'show_both_hazard_in_columns' => true
                        ];
                        
                        // ===== РЯД ДЛЯ SOC =====
                        
                        // Получаем стоимости для SOC
                        $socCost = $is40HC ? ceil(floatval($seaValue['SOC_40HC'] ?? 0)) : ceil(floatval($seaValue['SOC_20GP'] ?? 0));
                        
                        // Расчет для обычного груза SOC
                        $socNetto = ceil($dropOffCost + $socCost);
                        $costSeaSocNormal = ceil(($dropOffCost + $socNetto) * (1 + $cafPercent / 100) + $securityCostSea +  $seaProfit);

                        // Для SOC опасного используем те же стоимости что и для COC опасного
                        $costSeaSocDanger = $costSeaCocDanger;
                        
                        // Ж/Д часть для SOC (используем те же стоимости что и для COC)
                        $costRailSocNormal = $costRailCocNormal;
                        $costRailSocDanger = $costRailCocDanger;
                        
                        // Общие стоимости для SOC
                        $totalSocNormal = $costSeaSocNormal + $costRailSocNormal;
                        $totalSocDanger = $costSeaSocDanger + $costRailSocDanger;
                        
                        $resultItemSOC = [
                            'comb_sea_pol' => $seaValue['POL'] ?? '',
                            'comb_sea_pod' => $seaValue['POD'] ?? '',
                            'comb_rail_start' => $railStartStation,
                            'comb_rail_dest' => $railDestStation,
                            'comb_drop_off' => $seaValue['DROP_OFF_LOCATION'] ?? '',
                            'comb_transshipment_port' => $combValue['PUNKT_OTPRAVLENIYA'] ?? '',
                            'comb_coc' => $cocType,
                            'comb_container_ownership' => 'SOC',
                            'comb_hazard' => 'Оба варианта',
                            'comb_security' => $security === 'no' ? 'Нет' : ($security === '20' ? '20 фут' : '40 фут'),
                            'comb_agent' => trim(($seaValue['AGENT'] ?? '') . '; ' . ($railValue['AGENT'] ?? '')),
                            'comb_remark' => $this->getCombinedRemark($seaValue, $combValue, $railValue),
                            
                            // Морская часть - обычный груз
                            'cost_sea_normal' => $costSeaSocNormal,
                            'container_cost_normal' => $socCost,
                            'drop_off_cost_normal' => $dropOffCost,
                            'caf_percent_normal' => $cafPercent,
                            'security_cost_sea_normal' => $securityCostSea,
                            
                            // Морская часть - опасный груз
                            'cost_sea_danger' => $costSeaSocDanger,
                            'container_cost_danger' => $cocCostDanger,
                            'drop_off_cost_danger' => $dropOffCost,
                            'caf_percent_danger' => $cafPercent,
                            'security_cost_sea_danger' => $securityCostSea,
                            
                            // Ж/Д часть - обычный груз
                            'cost_rail_normal' => $costRailSocNormal,
                            'rail_base_cost_normal' => $railCostNormal,
                            'security_cost_rail_normal' => $securityCostRail,
                            
                            // Ж/Д часть - опасный груз
                            'cost_rail_danger' => $costRailSocDanger,
                            'rail_base_cost_danger' => $railCostDanger,
                            'security_cost_rail_danger' => $securityCostRail,
                            
                            // Общие стоимости
                            'cost_total_normal' => $totalSocNormal,
                            'cost_total_danger' => $totalSocDanger,
                            'cost_total_normal_text' => $costSeaSocNormal . '$ + ' . $costRailSocNormal . ' руб',
                            'cost_total_danger_text' => $costSeaSocDanger . '$ + ' . $costRailSocDanger . ' руб',
                            
                            'show_both_ownership' => false,
                            'show_both_hazard_in_columns' => true
                        ];
                        
                        // Добавляем оба ряда в результат
                        $result[] = $resultItemCOC;
                        $result[] = $resultItemSOC;
                        
                    } else {
                        // Если выбрана конкретная собственность контейнера
                        $selectedOwnership = $containerOwnership; // 'coc' или 'soc'
                        $displayContainerType = $selectedOwnership === 'coc' ? 'COC' : 'SOC';
                        
                        if (!$isHazard) {
                            // Пользователь выбрал обычный груз - показываем оба варианта (обычный и опасный)
                            
                            // Определяем тип контейнера для морской части
                            $is40HC = ($cocType === '40HC (28t)');
                            
                            // Получаем стоимости в зависимости от типа собственности
                            $containerCost = $selectedOwnership === 'soc' 
                                ? ($is40HC ? ceil(floatval($seaValue['SOC_40HC'] ?? 0)) : ceil(floatval($seaValue['SOC_20GP'] ?? 0)))
                                : ($is40HC ? ceil(floatval($seaValue['COC_40HC'] ?? 0)) : ceil(floatval($seaValue['COC_20GP'] ?? 0)));
                            
                            $containerCostDanger = $is40HC ? ceil(floatval($seaValue['OPASNYY_40HC'] ?? 0)) : ceil(floatval($seaValue['OPASNYY_20DC'] ?? 0));
                            
                            // DROP OFF стоимости
                            $dropOffCost = $is40HC ? ceil(floatval($seaValue['DROP_OFF_40HC'] ?? 0)) : ceil(floatval($seaValue['DROP_OFF_20GP'] ?? 0));
                            
                            // Стоимость охраны для морской части
                            $securityCostSea = 0;
                            if ($security === '20' && !$is40HC) {
                                $securityCostSea = ceil(floatval($seaValue['OKHRANA_20_FUT'] ?? 0));
                            } elseif ($security === '40' && $is40HC) {
                                $securityCostSea = ceil(floatval($seaValue['OKHRANA_40_FUT'] ?? 0));
                            }
                            
                            // CAF процент
                            $cafPercent = floatval($seaValue['CAF_KONVERT'] ?? 0);
                            
                            // Расчет для обычного груза
                            $netto = ceil($dropOffCost + $containerCost);
                            $costSeaNormal = ceil(($dropOffCost + $netto) * (1 + $cafPercent / 100) + $securityCostSea + $seaProfit);

                            // Расчет для опасного груза
                            $nettoDanger = ceil($dropOffCost + $containerCostDanger);
                            $costSeaDanger = ceil(($dropOffCost + $nettoDanger) * (1 + $cafPercent / 100) + $securityCostSea + $seaProfit);
                            
                            // ===== Ж/Д ЧАСТЬ =====
                            $railCostNormal = $this->getRailCostForContainerType($cocType, $railValue, false);
                            $railCostDanger = $this->getRailCostForContainerType($cocType, $railValue, true);
                            
                            // Стоимость охраны для ЖД части
                            $securityCostRail = $this->getSecurityCostForContainerType($railValue, $security, $cocType);

                            $costRailNormal = ceil($railCostNormal + $securityCostRail + $railProfit);
                            $costRailDanger = ceil($railCostDanger + $securityCostRail + $railProfit);
                            
                            // Общие стоимости
                            $totalNormal = $costSeaNormal + $costRailNormal;
                            $totalDanger = $costSeaDanger + $costRailDanger;
                            
                            $resultItem = [
                                'comb_sea_pol' => $seaValue['POL'] ?? '',
                                'comb_sea_pod' => $seaValue['POD'] ?? '',
                                'comb_rail_start' => $railStartStation,
                                'comb_rail_dest' => $railDestStation,
                                'comb_drop_off' => $seaValue['DROP_OFF_LOCATION'] ?? '',
                                'comb_transshipment_port' => $combValue['PUNKT_OTPRAVLENIYA'] ?? '',
                                'comb_coc' => $cocType,
                                'comb_container_ownership' => $displayContainerType,
                                'comb_hazard' => 'Нет',
                                'comb_security' => $security === 'no' ? 'Нет' : ($security === '20' ? '20 фут' : '40 фут'),
                                'comb_agent' => trim(($seaValue['AGENT'] ?? '') . '; ' . ($railValue['AGENT'] ?? '')),
                                'comb_remark' => $this->getCombinedRemark($seaValue, $combValue, $railValue),
                                
                                // Морская часть - обычный груз
                                'cost_sea_normal' => $costSeaNormal,
                                'container_cost_normal' => $containerCost,
                                'drop_off_cost_normal' => $dropOffCost,
                                'caf_percent_normal' => $cafPercent,
                                'security_cost_sea_normal' => $securityCostSea,
                                
                                // Морская часть - опасный груз
                                'cost_sea_danger' => $costSeaDanger,
                                'container_cost_danger' => $containerCostDanger,
                                'drop_off_cost_danger' => $dropOffCost,
                                'caf_percent_danger' => $cafPercent,
                                'security_cost_sea_danger' => $securityCostSea,
                                
                                // Ж/Д часть - обычный груз
                                'cost_rail_normal' => $costRailNormal,
                                'rail_base_cost_normal' => $railCostNormal,
                                'security_cost_rail_normal' => $securityCostRail,
                                
                                // Ж/Д часть - опасный груз
                                'cost_rail_danger' => $costRailDanger,
                                'rail_base_cost_danger' => $railCostDanger,
                                'security_cost_rail_danger' => $securityCostRail,
                                
                                // Общие стоимости
                                'cost_total_normal' => $totalNormal,
                                'cost_total_danger' => $totalDanger,
                                'cost_total_normal_text' => $costSeaNormal . '$ + ' . $costRailNormal . ' руб',
                                'cost_total_danger_text' => $costSeaDanger . '$ + ' . $costRailDanger . ' руб',
                                
                                'show_both_ownership' => false,
                                'show_both_hazard_in_columns' => true
                            ];
                            
                            $result[] = $resultItem;
                            
                        } else {
                            // Пользователь выбрал опасный груз - показываем только опасный
                            
                            // Определяем тип контейнера для морской части
                            $is40HC = ($cocType === '40HC (28t)');
                            
                            $containerCostDanger = $is40HC ? ceil(floatval($seaValue['OPASNYY_40HC'] ?? 0)) : ceil(floatval($seaValue['OPASNYY_20DC'] ?? 0));
                            
                            // DROP OFF стоимости
                            $dropOffCost = $is40HC ? ceil(floatval($seaValue['DROP_OFF_40HC'] ?? 0)) : ceil(floatval($seaValue['DROP_OFF_20GP'] ?? 0));
                            
                            // Стоимость охраны для морской части
                            $securityCostSea = 0;
                            if ($security === '20' && !$is40HC) {
                                $securityCostSea = ceil(floatval($seaValue['OKHRANA_20_FUT'] ?? 0));
                            } elseif ($security === '40' && $is40HC) {
                                $securityCostSea = ceil(floatval($seaValue['OKHRANA_40_FUT'] ?? 0));
                            }
                            
                            // CAF процент
                            $cafPercent = floatval($seaValue['CAF_KONVERT'] ?? 0);

                            // Расчет для опасного груза
                            $nettoDanger = ceil($dropOffCost + $containerCostDanger);
                            $costSeaDanger = ceil(($dropOffCost + $nettoDanger) * (1 + $cafPercent / 100) + $securityCostSea + $seaProfit);
                            
                            // ===== Ж/Д ЧАСТЬ ДЛЯ ОПАСНОГО ГРУЗА =====
                            $railCostDanger = $this->getRailCostForContainerType($cocType, $railValue, true);
                            
                            // Стоимость охраны для ЖД части
                            $securityCostRail = $this->getSecurityCostForContainerType($railValue, $security, $cocType);

                            $costRailDanger = ceil($railCostDanger + $securityCostRail + $railProfit);
                            
                            // Общая стоимость
                            $totalDanger = $costSeaDanger + $costRailDanger;
                            
                            $resultItem = [
                                'comb_sea_pol' => $seaValue['POL'] ?? '',
                                'comb_sea_pod' => $seaValue['POD'] ?? '',
                                'comb_rail_start' => $railStartStation,
                                'comb_rail_dest' => $railDestStation,
                                'comb_drop_off' => $seaValue['DROP_OFF_LOCATION'] ?? '',
                                'comb_transshipment_port' => $combValue['PUNKT_OTPRAVLENIYA'] ?? '',
                                'comb_coc' => $cocType,
                                'comb_container_ownership' => $displayContainerType,
                                'comb_hazard' => 'Да',
                                'comb_security' => $security === 'no' ? 'Нет' : ($security === '20' ? '20 фут' : '40 фут'),
                                'comb_agent' => trim(($seaValue['AGENT'] ?? '') . '; ' . ($railValue['AGENT'] ?? '')),
                                'comb_remark' => $this->getCombinedRemark($seaValue, $combValue, $railValue),
                                
                                // Морская часть - опасный груз
                                'cost_sea' => $costSeaDanger,
                                'container_cost' => $containerCostDanger,
                                'drop_off_cost' => $dropOffCost,
                                'caf_percent' => $cafPercent,
                                'security_cost_sea' => $securityCostSea,
                                
                                // Ж/Д часть - опасный груз
                                'cost_rail' => $costRailDanger,
                                'rail_base_cost' => $railCostDanger,
                                'security_cost_rail' => $securityCostRail,
                                
                                // Общая стоимость
                                'cost_total' => $totalDanger,
                                'cost_total_text' => $costSeaDanger . '$ + ' . $costRailDanger . ' руб',
                                
                                'show_both_ownership' => false,
                                'show_both_hazard_in_columns' => false
                            ];
                            
                            $result[] = $resultItem;
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
        $remarks[] = 'Море: ' . trim($seaValue['REMARK']);
    }
    
    // Добавляем примечание из комбинированной перевозки
    if (!empty($combPerevozki)) {
        foreach ($combPerevozki as $combItem) {
            if (trim($combItem['STANTSIYA_OTPRAVLENIYA'] ?? '') === trim($railStartStation)) {
                if (!empty($combItem['REMARK'])) {
                    $remarks[] = 'Комб: ' . trim($combItem['REMARK']);
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
                            'PROPERTY_166' => str_replace(',', '', trim((string)$row['C'])),
                            'PROPERTY_168' => str_replace(',', '', trim((string)$row['D'])),
                            'PROPERTY_170' => str_replace(',', '', trim((string)$row['E'])),
                            'PROPERTY_172' => str_replace(',', '', trim((string)$row['F'])),
                            'PROPERTY_174' => str_replace(',', '', trim((string)$row['G'])),
                            'PROPERTY_176' => str_replace(',', '', trim((string)$row['H'])),
                            'PROPERTY_178' => str_replace(',', '', trim((string)$row['I'])),
                            'PROPERTY_180' => str_replace(',', '', trim((string)$row['J'])),
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
                        'FIELDS' => [
                            'NAME'         => trim((string)$row['A']),  // Порт
                            'PROPERTY_126' => trim((string)$row['B']),
                            'PROPERTY_162' => str_replace(',', '', trim((string)$row['C'])),
                            'PROPERTY_164' => str_replace(',', '', trim((string)$row['D'])),
                            'PROPERTY_132' => trim((string)$row['E']),
                            'PROPERTY_134' => str_replace(',', '', trim((string)$row['F'])),
                            'PROPERTY_136' => str_replace(',', '', trim((string)$row['G'])),
                            'PROPERTY_138' => trim((string)$row['H']),
                            'PROPERTY_140' => trim((string)$row['I']),
                            'PROPERTY_192' => trim((string)$row['J']), // agent
                            'PROPERTY_200' => str_replace(',', '', trim((string)$row['K'])),
                            'PROPERTY_202' => str_replace(',', '', trim((string)$row['L'])),
                            'PROPERTY_208' => str_replace(',', '', trim((string)$row['M'])),
                            'PROPERTY_210' => str_replace(',', '', trim((string)$row['N'])),
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