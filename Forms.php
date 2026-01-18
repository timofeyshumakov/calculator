<?php
/**
 * Форма калькулятора расчета стоимости перевозок (Bootstrap 5) с зависимым списком станций Ж/Д перевозок
 * 
 * @var array $seaPerevozki
 * @var array $zhdPerevozki
 * @var array $combPerevozki
 */

// Списки морских портов
$seaPorts = array_unique(array_column($seaPerevozki, 'POL'));

// Списки Ж/Д станций отправлений
$zhdStarts = array_unique(array_column($zhdPerevozki, 'POL'));

// Списки комбинированных портов отправлений
$combStarts = array_unique(array_column($seaPerevozki, 'POL'));

?>
<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Калькулятор стоимости перевозок</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.10.5/font/bootstrap-icons.css" />
    <style>
        .calc-section { display: none; }
        .hiden {
            display: none;
        }
        /* Стили для индикатора загрузки */
        .loading-spinner {
            border: 3px solid #f3f3f3;
            border-top: 3px solid #3498db;
            border-radius: 50%;
            width: 40px;
            height: 40px;
            animation: spin 1s linear infinite;
        }
        
        @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
        }
        
        /* Стили для прогресс-бара */
        .progress-bar {
            transition: width 0.3s ease;
        }
        
        /* Стили для кнопок экспорта */
        .export-btn {
            transition: all 0.3s ease;
        }
        
        .export-btn:hover {
            transform: translateY(-2px);
            box-shadow: 0 4px 8px rgba(0,0,0,0.1);
        }
    </style>
</head>
<body>
<div class="container-fluid py-4 px-4">
    <!-- Индикатор загрузки (изначально скрыт) -->
    <div id="loading-overlay" style="display: none; position: fixed; top: 0; left: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.5); z-index: 9999;">
        <div class="d-flex justify-content-center align-items-center h-100">
            <div class="text-center bg-white p-4 rounded shadow">
                <div class="spinner-border text-primary mb-3" style="width: 3rem; height: 3rem;" role="status">
                    <span class="visually-hidden">Загрузка...</span>
                </div>
                <h5 class="mb-2">Идет загрузка...</h5>
                <p id="loading-message" class="text-muted mb-0">Пожалуйста, подождите</p>
                <div class="progress mt-3" style="height: 6px;">
                    <div id="loading-progress" class="progress-bar progress-bar-striped progress-bar-animated" 
                         style="width: 0%;"></div>
                </div>
            </div>
        </div>
    </div>
    
    <!-- Блок для сообщений о загрузке файлов -->
    <div id="upload-status" class="alert alert-dismissible fade show mb-4" style="display: none;" role="alert">
        <span id="upload-status-text"></span>
        <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button>
    </div>
    <div class="d-flex justify-content-between align-items-center mb-4">
        <h1 class="mb-0">Калькулятор стоимости перевозок</h1>
        <div>
            <button type="button" id="upload_comb" class="btn btn-primary me-2">
            <i class="bi bi-upload"></i> Комбинированный
            </button>
            <button type="button" id="upload_sea" class="btn btn-primary me-2">
            <i class="bi bi-upload"></i> Морские
            </button>
            <button type="button" id="upload_zhd" class="btn btn-primary">
            <i class="bi bi-upload"></i> Ж/Д
            </button>
        </div>
    </div>
    <form action="calculate.php" method="post">
        <div class="mb-4">
            <label for="calc_type" class="form-label">Выберите тип расчёта</label>
            <select id="calc_type" name="calc_type" class="form-select" required>
                <option value="">-- Выберите --</option>
                <option value="sea">Расчёт морских перевозок</option>
                <option value="rail">Расчёт ж/д перевозок</option>
                <option value="combined">Комбинированный маршрут</option>
            </select>
        </div>

        <!-- МОРСКИЕ ПЕРЕВОЗКИ -->
        <div id="section_sea" class="card mb-4 calc-section">
            <div class="card-header d-flex justify-content-between align-items-center">
                <span>Морские перевозки</span>
            </div>
            <div class="card-body">
                <div class="row g-3">
                    <div class="col-md-6">
                        <label for="sea_pol" class="form-label">POL (Порт отправления)</label>
                        <select id="sea_pol" name="sea_pol" class="form-select">
                            <option value="">Выберите порт отправления...</option>
                            <?php foreach ($seaPorts as $port): ?>
                            <option value="<?= htmlspecialchars($port) ?>">
                                <?= htmlspecialchars($port) ?>
                            </option>
                            <?php endforeach; ?>
                        </select>
                    </div>
                    <div class="col-md-6">
                        <label for="sea_pod" class="form-label">POD (Порт прибытия)</label>
                        <select id="sea_pod" name="sea_pod" class="form-select" disabled>
                            <option value="">Выберите порт прибытия...</option>
                        </select>
                    </div>
                    <div class="col-md-3">
                        <label for="sea_coc" class="form-label">ТИП КОНТЕЙНЕРА</label>
                        <select id="sea_coc" name="sea_coc" class="form-select" disabled>
                            <option value="">Выберите тип...</option>
                        </select>
                    </div>
                    <div class="col-md-2">
                        <label for="sea_hazard" class="form-label">Опасный груз?</label>
                        <select id="sea_hazard" name="sea_hazard" class="form-select">
                            <option value="no">Нет</option>
                            <option value="yes">Да</option>
                        </select>
                    </div>
                    <div class="col-md-3">
                        <label for="sea_security" class="form-label">Охрана</label>
                        <select id="sea_security" name="sea_security" class="form-select">
                            <option value="no">Нет</option>
                            <option value="20">20 фут</option>
                            <option value="40">40 фут</option>
                        </select>
                    </div>
                    <div class="col-md-3">
                        <label for="sea_container_ownership" class="form-label">СОБСТВЕННОСТЬ КОНТЕЙНЕРА</label>
                        <select id="sea_container_ownership" name="sea_container_ownership" class="form-select">
                            <option value="no">Не выбрано</option>
                            <option value="coc">COC</option>
                            <option value="soc">SOC</option>
                        </select>
                    </div>
                    <div class="col-md-3 hiden">
                        <label for="sea_agent" class="form-label">Агент</label>
                        <input type="text" id="sea_agent" name="sea_agent" class="form-control" placeholder="Агент" disabled />
                    </div>
                    <div class="col-md-6 hiden">
                        <label for="sea_remark" class="form-label">Remark</label>
                        <input type="text" id="sea_remark" name="sea_remark" class="form-control" placeholder="Комментарий" />
                    </div>
                    <div class="col-md-6">
                        <label for="sea_drop_off_location" class="form-label">DROP OFF LOCATION</label>
                        <select id="sea_drop_off_location" name="sea_drop_off_location" class="form-select" disabled>
                            <option value="">Выберите...</option>
                        </select>
                    </div>
                    <div class="col-md-6 hiden">
                        <label for="sea_drop_off" class="form-label">DROP OFF (Число)</label>
                        <input type="number" id="sea_drop_off" name="sea_drop_off" class="form-control" />
                    </div>
                    <div class="col-md-4">
                        <label for="sea_caf" class="form-label">% CAF (конверт)</label>
                        <input type="number" step="0.5" id="sea_caf" name="sea_caf" class="form-control" placeholder="%" />
                    </div>
                    <div class="col-md-4 hiden">
                        <label for="sea_netto" class="form-label">NETTO (COST) USD</label>
                        <input type="number" step="1" id="sea_netto" name="sea_netto" class="form-control" placeholder="0" />
                    </div>
                    <div class="col-md-4">
                        <label for="sea_profit" class="form-label">Profit (Море, $)</label>
                        <input type="number" step="1" id="sea_profit" name="sea_profit" class="form-control" placeholder="0" />
                    </div>
                    <div class="col-md-12 hiden">
                        <label for="sea_brutto" class="form-label">БРУТТО (PRICE) USD</label>
                        <input type="number" step="1" id="sea_brutto" name="sea_brutto" class="form-control" readonly placeholder="Вычисляется автоматически" disabled />
                    </div>
                    <div class="col-md-12">
                        <div class="d-flex gap-2">
                            <button type="button" id="sea_calculate" class="btn btn-primary disabled">Рассчитать</button>
                        </div>
                    </div>
                    <div class="sea_result mt-4"></div>
                </div>
            </div>
        </div>

        <!-- ПОЕЗД -->
        <div id="section_rail" class="card mb-4 calc-section">
            <div class="card-header d-flex justify-content-between align-items-center">
                <span>Ж/Д перевозки</span>
            </div>
            <div class="card-body">
                <div class="row g-3">
                    <div class="col-md-6">
                        <label for="rail_origin" class="form-label">Станция отправления</label>
                        <select id="rail_origin" name="rail_origin" class="form-select">
                            <option value="">Выберите станцию отправления...</option>
                            <?php foreach ($zhdStarts as $station): ?>
                                <option value="<?= htmlspecialchars($station) ?>">
                                    <?= htmlspecialchars($station) ?>
                                </option>
                            <?php endforeach; ?>
                        </select>
                    </div>
                    <div class="col-md-6">
                        <label for="rail_destination" class="form-label">Станция назначения</label>
                        <select id="rail_destination" name="rail_destination" class="form-select">
                            <option value="">Выберите пункт...</option>
                        </select>
                    </div>
                    <div class="col-md-4">
                        <label for="rail_coc" class="form-label">ТИП КОНТЕЙНЕРА</label>
                        <select id="rail_coc" name="rail_coc" class="form-select">
                            <option value="">Выберите...</option>
                            <option>20DC (<24t)</option>
                            <option>20DC (24t-28t)</option>
                            <option>40HC (28t)</option>
                        </select>
                    </div>
                    <div class="col-md-3">
                        <label for="rail_container_ownership" class="form-label">СОБСТВЕННОСТЬ КОНТЕЙНЕРА</label>
                        <select id="rail_container_ownership" name="rail_container_ownership" class="form-select">
                            <option value="no">Не выбрано</option>
                            <option value="coc">COC</option>
                            <option value="soc">SOC</option>
                        </select>
                    </div>
                    <div class="col-md-3 hiden">
                        <label for="rail_agent" class="form-label">Агент</label>
                        <input type="text" id="rail_agent" name="rail_agent" class="form-control" placeholder="Агент" disabled />
                    </div>
                    <div class="col-md-2 hiden">
                        <label for="rail_hazard" class="form-label">Опасный груз?</label>
                        <select id="rail_hazard" name="rail_hazard" class="form-select">
                            <option value="no">Нет</option>
                            <option value="yes">Да</option>
                        </select>
                    </div>
                    <div class="col-md-3">
                        <label for="rail_security" class="form-label">Охрана</label>
                        <select id="rail_security" name="rail_security" class="form-select">
                            <option value="no">Нет</option>
                            <option value="20">20 фут</option>
                            <option value="40">40 фут</option>
                        </select>
                    </div>
                    <div class="col-md-4">
                        <label for="rail_profit" class="form-label">Profit (ЖД, ₽)</label>
                        <input type="number" step="1" id="rail_profit" name="rail_profit" class="form-control" placeholder="0" />
                    </div>
                    <div class="col-md-12 hiden">
                        <label for="rail_sum" class="form-label">ИТОГО</label>
                        <input type="number" step="1" id="rail_sum" name="rail_sum" class="form-control" readonly placeholder="Вычисляется автоматически" disabled />
                    </div>
                    <div class="col-md-12">
                        <div class="d-flex gap-2">
                            <button type="button" id="rail_calculate" class="btn btn-primary disabled">Рассчитать</button>
                        </div>
                    </div>
                    <div class="rail_result mt-4"></div>
                </div>
            </div>
        </div>

        <!-- КОМБИНИРОВАННЫЙ -->
        <div id="section_combined" class="card mb-4 calc-section">
            <div class="card-header d-flex justify-content-between align-items-center">
                <span>Комбинированный маршрут</span>
            </div> 
            <div class="card-body">
                <div class="row g-3">
                    <div class="col-md-4">
                        <label for="comb_sea_pol" class="form-label">POL (Порт отправления)</label>
                        <select id="comb_sea_pol" name="comb_sea_pol" class="form-select">
                            <option value="">Выберите порт...</option>
                            <?php foreach ($combStarts as $port): ?>
                            <option value="<?= htmlspecialchars($port) ?>">
                                <?= htmlspecialchars($port) ?>
                            </option>
                            <?php endforeach; ?>
                        </select>
                    </div>
                    <div class="col-md-4">
                        <label for="comb_drop_off" class="form-label">DROP OFF</label>
                        <select id="comb_drop_off" name="comb_drop_off" class="form-select" disabled>
                            <option value="">Выберите станцию...</option>
                        </select>
                    </div>
                    <div class="col-md-4">
                        <label for="comb_coc" class="form-label">ТИП КОНТЕЙНЕРА</label>
                        <select id="comb_coc" name="comb_coc" class="form-select">
                            <option value="">Выберите тип...</option>
                        </select>
                    </div>
                    <div class="col-md-3">
                        <label for="comb_container_ownership" class="form-label">СОБСТВЕННОСТЬ КОНТЕЙНЕРА</label>
                        <select id="comb_container_ownership" name="comb_container_ownership" class="form-select">
                            <option value="no">Не выбрано</option>
                            <option value="coc">COC</option>
                            <option value="soc">SOC</option>
                        </select>
                    </div>
                    <div class="col-md-4">
                        <label for="comb_hazard" class="form-label">Опасный груз?</label>
                        <select id="comb_hazard" name="comb_hazard" class="form-select">
                            <option value="no">Нет</option>
                            <option value="yes">Да</option>
                        </select>
                    </div>
                    <div class="col-md-5">
                        <label for="comb_transshipment_port" class="form-label">ПОРТ ПЕРЕВАЛКИ</label>
                        <select id="comb_transshipment_port" name="comb_transshipment_port" class="form-select">
                        </select>
                    </div>
                    <div class="col-md-8 hiden">
                        <label for="comb_remark" class="form-label">Remark</label>
                        <input type="text" id="comb_remark" name="comb_remark" class="form-control" placeholder="Комментарий" disabled />
                    </div>
                    <div class="col-md-6">
                        <label for="comb_security" class="form-label">Охрана</label>
                        <select id="comb_security" name="comb_security" class="form-select">
                            <option value="Нет" selected>Нет</option>
                            <option>20 фут</option>
                            <option>40 фут</option>
                        </select>
                    </div>
                    <div class="col-md-4">
                        <label for="sea_profit" class="form-label">Profit (Море, $)</label>
                        <input type="number" step="1" id="sea_profit" name="sea_profit" class="form-control" placeholder="0" />
                    </div>
                    <div class="col-md-4">
                        <label for="rail_profit" class="form-label">Profit (ЖД, ₽)</label>
                        <input type="number" step="1" id="rail_profit" name="rail_profit" class="form-control" placeholder="0" />
                    </div>
                    <div class="col-md-12">
                        <div class="d-flex gap-2">
                            <button type="button" id="comb_calculate" class="btn btn-primary disabled">Рассчитать</button>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <!-- КОМБИНИРОВАННЫЙ РЕЗУЛЬТАТ -->
        <div class="comb_result"></div>

        <!-- INPUT ДЛЯ ЗАГРУЗКИ ФАЙЛА -->
        <input type="file" id="upload_file_input" accept=".xlsx" style="display: none;" />
    </form>
</div>

<!-- Bootstrap 5 JS -->
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
<!-- Скрипт для показа нужного раздела и расчета BRUTTO -->
<script>
    const seaPerevozki = <?= json_encode($seaPerevozki, JSON_UNESCAPED_UNICODE) ?>;
    const cleanUrl = window.location.origin + window.location.pathname;
    // Глобальные переменные для хранения результатов
    let currentSeaResults = [];
    let currentRailResults = [];
    let currentCombResults = [];

    document.addEventListener('DOMContentLoaded', function() {
        function getLocationMapping() {
            return {
                'MOSCOW': 'Москва',
                'ST PETERBURG': 'Санкт-Петербург',
                'NOVOSIBIRSK': 'Новосибирск',
                'EKATERINBURG': 'Екатеринбург',
                'VLADIVOSTOK': 'Владивосток',
                'SAMARA': 'Самара',
                'TOLYATTI': 'Тольятти',
                'KRASNOYARSK': 'Красноярск',
                'IRKUTSK': 'Иркутск',
                'VRANGEL BAY': 'Порт Врангель',
            };
        }
        // Секции форм
        const selector = document.getElementById('calc_type');
        const sections = {
            sea: document.getElementById('section_sea'),
            rail: document.getElementById('section_rail'),
            combined: document.getElementById('section_combined')
        };
        function updateSections() {
            const val = selector.value;
            Object.keys(sections).forEach(key => {
                sections[key].style.display = (key === val ? 'block' : 'none');
            });
            if (val !== 'combined') {
                const combResult = document.querySelector('.comb_result');
                if (combResult) {
                    combResult.innerHTML = ''; // Очищаем содержимое
                }
            }
        }
        selector.addEventListener('change', updateSections);
        updateSections();

        // МОРСКИЕ ПЕРЕВОЗКИ
        // Преобразуем PHP-массив seaPerevozki в JS
        const seaPerevozki = <?php echo json_encode($seaPerevozki, JSON_UNESCAPED_UNICODE); ?>;

        const polSelect   = document.getElementById('sea_pol');
        const dropSelect = document.getElementById('comb_drop_off');
        const combPolSelect = document.getElementById('comb_sea_pol');
        const podSelect   = document.getElementById('sea_pod');
        const cocSelect   = document.getElementById('sea_coc');
        const dropLocSel  = document.getElementById('sea_drop_off_location');
        const dropOffInp  = document.getElementById('sea_drop_off');
        const cafInp      = document.getElementById('sea_caf');
        const remarkInp   = document.getElementById('sea_remark');
        const agentInp   = document.getElementById('sea_agent');
        const bruttoInp   = document.getElementById('sea_brutto');
        const nettoInp    = document.getElementById('sea_netto');
        const profitInp   = document.getElementById('sea_profit');
        const seaOwnership   = document.getElementById('sea_container_ownership');

        function reset(select, placeholder) {
            select.innerHTML = `<option value="">${placeholder}</option>`;
            select.disabled = true;
        }

        // При выборе POL — наполняем POD
        polSelect.addEventListener('change', function() {
            reset(podSelect, 'Выберите порт прибытия...');
            reset(dropLocSel, 'Выберите терминал...');
            resetCocAndFields();
            const pol = this.value;
            if (!pol) return;

            // Собираем уникальные POD для выбранного POL
            const pods = [...new Set(seaPerevozki
                .filter(r => r.POL === pol)
                .map(r => r.POD))];
            pods.forEach(pod => {
                const opt = document.createElement('option');
                opt.value = pod;
                opt.text  = pod;
                podSelect.appendChild(opt);
            });
            podSelect.disabled = false;
        });

        combPolSelect.addEventListener('change', function () {
            const selectedPol = this.value;

            // очищаем список
            dropSelect.innerHTML = '<option value="">Выберите...</option>';

            if (!selectedPol) {
                dropSelect.disabled = true;
                return;
            }

            // выбираем все элементы с этим POL
            const matches = seaPerevozki.filter(item => item.POL === selectedPol);
            // получаем уникальные DROP_OFF_LOCATION
            const dropOffs = [...new Set(matches.map(item => item.DROP_OFF_LOCATION))];

            // наполняем select
            dropOffs.forEach(loc => {
                const opt = document.createElement('option');
                opt.value = loc;
                opt.textContent = loc;
                dropSelect.appendChild(opt);
            });

            dropSelect.disabled = dropOffs.length === 0;
        });
        // При выборе POD — наполняем DROP OFF LOCATION
        podSelect.addEventListener('change', function() {
            reset(dropLocSel, 'Выберите терминал...');
            resetCocAndFields();
            const pol = polSelect.value;
            const pod = this.value;
            if (!pod) return;

            // Уникальные DROP_OFF_LOCATION
            const locs = [...new Set(seaPerevozki
                .filter(r => r.POL === pol && r.POD === pod)
                .map(r => r.DROP_OFF_LOCATION))];
            locs.forEach(loc => {
                const opt = document.createElement('option');
                opt.value = loc;
                opt.text  = loc;
                dropLocSel.appendChild(opt);
            });
            dropLocSel.disabled = false;
        });

        // При выборе DROP_OFF_LOCATION — заполняем COC, CAF, Remark
        dropLocSel.addEventListener('change', function() {
            reset(cocSelect, 'Выберите тип контейнера...');
            cafInp.value = agentInp.value = remarkInp.value = dropOffInp.value = '';
            bruttoInp.value = profitInp.value = '';

            const pol = polSelect.value;
            const pod = podSelect.value;
            const loc = this.value;
            if (!(pol && pod && loc)) return;

            // Берём первую запись для пол+под+терминал
            const rec = seaPerevozki.find(r =>
                r.POL === pol &&
                r.POD === pod &&
                r.DROP_OFF_LOCATION === loc
            );
            if (!rec) return;

            // Заполняем селект COC двумя вариантами
            cocSelect.innerHTML = '<option value="">Выберите тип контейнера...</option>';
            ['20GP', '40HC'].forEach(type => {
                const opt = document.createElement('option');
                opt.value = type;
                opt.text  = type;
                cocSelect.appendChild(opt);
            });
            cocSelect.disabled = false;

            // CAF и Remark
            cafInp.value    = rec.CAF_KONVERT;
            remarkInp.value = rec.REMARK ?? '';
            agentInp.value  = rec.AGENT ?? '';
        });
        // Пересчёт drop_off и NETTO+BRUTTO
        cocSelect.addEventListener('change', updateCosts);
        nettoInp.addEventListener('input', updateBrutto);
        profitInp.addEventListener('input', updateBrutto);
        // добавляем для CAF
        cafInp.addEventListener('input', updateCosts);
        // (опционально) для ручного правки DROP OFF
        dropOffInp.addEventListener('input', updateCosts);

        function updateCosts() {
            // Сбрасывать только поля расчёта, а не селекты, чтобы не мешать повторным вызовам:
            bruttoInp.value = nettoInp.value  = '';

            const pol = polSelect.value;
            const pod = podSelect.value;
            const loc = dropLocSel.value;
            const coc = cocSelect.value;
            // Если контейнер ещё не выбран — просто пересчитаем BRUTTO по новым profit/netto:
            if (!(pol && pod && loc && coc)) {
                return updateBrutto();
            }

            // находим запись и достаём числа
            const rec = seaPerevozki.find(r =>
                r.POL === pol &&
                r.POD === pod &&
                r.DROP_OFF_LOCATION === loc
            );
            if (!rec) return updateBrutto();
            const costCOC = parseFloat(coc === '20GP' ? rec[`${seaOwnership.value.toUpperCase()}_20GP`] : rec[`${seaOwnership.value.toUpperCase()}_40HC`]);
            const dropOff = parseFloat(coc === '20GP' ? rec.DROP_OFF_20GP : rec.DROP_OFF_40HC);
            const caf     = parseFloat(cafInp.value) || 0;

            dropOffInp.value = dropOff;
            const netto = (costCOC + dropOff);
            nettoInp.value   = netto;

            updateBrutto();
        }

        function updateBrutto() {
            const netto  = parseFloat(nettoInp.value)  || 0;
            const profit = parseFloat(profitInp.value) || 0;
            const caf = parseFloat(cafInp.value) || 0;
            const dropOff = parseFloat(dropOffInp.value) || 0;
            bruttoInp.value = ((dropOff + netto) * (caf / 100 + 1) + profit);
        }

        function resetCocAndFields() {
            reset(cocSelect, 'Выберите тип...');
            cafInp.value = remarkInp.value  = '';//= dropOffInp.value = bruttoInp.value
        }
        // ===== РАСЧЕТ МОРСКИХ ПЕРЕВОЗОК =====
        const seaCalculateBtn = document.getElementById('sea_calculate');
        seaCalculateBtn.addEventListener('click', function() {
            if (this.classList.contains('disabled')) return;
            showLoading('Расчет морских перевозок...');
            // Собираем данные формы
            const payload = {
                sea_pol: document.getElementById('sea_pol').value,
                sea_pod: document.getElementById('sea_pod').value,
                sea_drop_off_location: document.getElementById('sea_drop_off_location').value,
                sea_coc: document.getElementById('sea_coc').value,
                sea_container_ownership: document.getElementById('sea_container_ownership').value,
                sea_hazard: document.getElementById('sea_hazard').value,
                sea_security: document.getElementById('sea_security').value,
                sea_caf: document.getElementById('sea_caf').value,
                sea_profit: document.getElementById('sea_profit').value
            };
            
            // Валидация
            if (!payload.sea_pol || !payload.sea_pod) {
                alert('Пожалуйста, заполните порт отправления и порт прибытия');
                return;
            }
            
            // Отправка запроса
            fetch(cleanUrl + '?action=getSeaPerevozki', {
                method: 'POST',
                headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                body: new URLSearchParams(payload)
            })
            .then(res => res.json())
            .then(data => {
                if (data.error) {
                    alert(data.message || 'Ошибка расчета');
                    return;
                }
                
                // Сохраняем результаты для экспорта
                currentSeaResults = data;
                
                // Показываем кнопку экспорта
                const exportBtn = document.getElementById('export_sea');
                if (exportBtn) {
                    exportBtn.style.display = 'inline-block';
                }
                
                const seaResult = document.querySelector('.sea_result');
                if (!seaResult) {
                    const seaSection = document.getElementById('section_sea');
                    const resultDiv = document.createElement('div');
                    resultDiv.className = 'sea_result';
                    seaSection.appendChild(resultDiv);
                }
                
                let html = `
                <div class="card mt-4">
                    <div class="card-header d-flex justify-content-between align-items-center">
                        <h4 class="mb-0">Результаты расчета морских перевозок</h4>
                        <div>
                            <button type="button" id="export_sea_table" class="btn btn-success btn-sm export-btn">
                                <i class="bi bi-file-excel"></i> Экспорт в Excel
                            </button>
                        </div>
                    </div>
                    <div class="card-body">
                `;
                if (data.length === 0) {
                    html += '<div class="alert alert-info">Нет данных для выбранных параметров</div>';
                } else {
                    html += `
                        <div class="table-responsive">
                            <table class="table table-bordered table-striped">
                                <thead class="table-light">
                                    <tr>
                                        <th>Порт отправления</th>
                                        <th>Порт прибытия</th>
                                        <th>DROP OFF LOCATION</th>
                                        <th>Тип контейнера</th>
                                        <th>Собственность контейнера</th>
                                        <th>Опасный груз</th>
                                        <th>Охрана</th>
                                        <th>Стоимость контейнера ($)</th>
                                        <th>Стоимость DROP OFF ($)</th>
                                        <th>Стоимость охраны ($)</th>
                                        <th>NETTO ($)</th>
                                        <th>CAF (%)</th>
                                        <th>Profit ($)</th>
                                        <th>Итоговая стоимость ($)</th>
                                        <th>Итоговая стоимость (опасный) ($)</th>
                                        <th>Агент</th>
                                        <th>Примечание</th>
                                    </tr>
                                </thead>
                                <tbody>
                    `;
                    
                    data.forEach(item => {
                        html += `
                            <tr>
                                <td>${item.sea_pol || ''}</td>
                                <td>${item.sea_pod || ''}</td>
                                <td>${item.sea_drop_off_location || ''}</td>
                                <td>${item.sea_coc || ''}</td>
                                <td>${item.sea_container_ownership || ''}</td>
                                <td>${item.sea_hazard || 'Нет'}</td>
                                <td>${item.sea_security || 'Нет'}</td>
                                <td>${item.cost_container || 0}</td>
                                <td>${item.cost_drop_off || 0}</td>
                                <td>${item.cost_security || 0}</td>
                                <td>${item.cost_netto || 0}</td>
                                <td>${item.sea_caf_percent || 0}</td>
                                <td>${item.sea_profit || 0}</td>
                                <td><strong>${item.cost_total || 0}</strong></td>
                                <td><strong>${item.cost_total_danger || item.cost_total || 0}</strong></td>
                                <td>${item.sea_agent || ''}</td>
                                <td>${item.sea_remark || ''}</td>
                            </tr>`;
                    });
                    
                    html += '</tbody></table></div>';
                }
                
                html += '</div></div>';
                document.querySelector('.sea_result').innerHTML = html;
                
                // Добавляем обработчик для кнопки экспорта в таблице
                const exportTableBtn = document.getElementById('export_sea_table');
                if (exportTableBtn) {
                    exportTableBtn.addEventListener('click', () => {
                        exportToExcel('sea');
                    });
                }
                
                hideLoading();
            })
            .catch(err => {
                console.error(err);
                hideLoading();
                alert('Ошибка при расчете');
            });
        });

        // ===== РАСЧЕТ Ж/Д ПЕРЕВОЗОК =====
        const railCalculateBtn = document.getElementById('rail_calculate');
        railCalculateBtn.addEventListener('click', function() {
            if (this.classList.contains('disabled')) return;
            showLoading('Расчет ж/д перевозок...');
            // Собираем данные формы
            const payload = {
                rail_origin: document.getElementById('rail_origin').value,
                rail_destination: document.getElementById('rail_destination').value,
                rail_coc: document.getElementById('rail_coc').value,
                rail_container_ownership: document.getElementById('rail_container_ownership').value,
                rail_hazard: document.getElementById('rail_hazard').value,
                rail_security: document.getElementById('rail_security').value,
                rail_profit: document.getElementById('rail_profit').value
            };
            
            // Валидация
            if (!payload.rail_origin || !payload.rail_destination) {
                alert('Пожалуйста, заполните станцию отправления и станцию назначения');
                return;
            }
            
            // Отправка запроса
            fetch(cleanUrl + '?action=getRailPerevozki', {
                method: 'POST',
                headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                body: new URLSearchParams(payload)
            })
            .then(res => res.json())
            .then(data => {
                if (data.error) {
                    alert(data.message || 'Ошибка расчета');
                    return;
                }
                
                // Сохраняем результаты для экспорта
                currentRailResults = data;
                
                // Показываем кнопку экспорта
                const exportBtn = document.getElementById('export_rail');
                if (exportBtn) {
                    exportBtn.style.display = 'inline-block';
                }
                
                const railResult = document.querySelector('.rail_result');
                if (!railResult) {
                    const railSection = document.getElementById('section_rail');
                    const resultDiv = document.createElement('div');
                    resultDiv.className = 'rail_result';
                    railSection.appendChild(resultDiv);
                }
                
                let html = `
                <div class="card mt-4">
                    <div class="card-header d-flex justify-content-between align-items-center">
                        <h4 class="mb-0">Результаты расчета Ж/Д перевозок</h4>
                        <div>
                            <button type="button" id="export_rail_table" class="btn btn-success btn-sm export-btn">
                                <i class="bi bi-file-excel"></i> Экспорт в Excel
                            </button>
                        </div>
                    </div>
                    <div class="card-body">
                `;

                if (data.length === 0) {
                    html += '<div class="alert alert-info">Нет данных для выбранных параметров</div>';
                } else {
                    html += `
                        <div class="table-responsive">
                            <table class="table table-bordered table-striped">
                                <thead class="table-light">
                                    <tr>
                                        <th>Станция отправления</th>
                                        <th>Пункт назначения</th>
                                        <th>Тип контейнера</th>
                                        <th>Собственность контейнера</th>
                                        <th>Опасный груз</th>
                                        <th>Охрана</th>
                                        <th>Profit (₽)</th>
                                        <th>Базовая стоимость 20DC (<24t) (₽)</th>
                                        <th>Базовая стоимость 20DC (24t-28t) (₽)</th>
                                        <th>Базовая стоимость 40HC (28t) (₽)</th>
                                        <th>Стоимость охраны (₽)</th>
                                        <th>Итого 20DC (<24t) (₽)</th>
                                        <th>Итого 20DC (24t-28t) (₽)</th>
                                        <th>Итого 40HC (28t) (₽)</th>
                                        <th>Агент</th>
                                    </tr>
                                </thead>
                                <tbody>
                    `;
                    
                    data.forEach(item => {
                        html += `
                            <tr>
                                <td>${item.rail_origin || ''}</td>
                                <td>${item.rail_destination || ''}</td>
                                <td>${item.rail_coc || ''}</td>
                                <td>${item.rail_container_ownership || ''}</td>
                                <td>${item.rail_hazard || ''}</td>
                                <td>${item.rail_security || ''}</td>
                                <td>${item.rail_profit || 0}</td>
                                <td>${item.cost_base_20 || 0}</td>
                                <td>${item.cost_base_20_28 || 0}</td>
                                <td>${item.cost_base_40 || 0}</td>
                                <td>${item.cost_security || 0}</td>
                                <td><strong>${item.cost_total_20 || 0}</strong></td>
                                <td><strong>${item.cost_total_20_28 || 0}</strong></td>
                                <td><strong>${item.cost_total_40 || 0}</strong></td>
                                <td>${item.rail_agent || ''}</td>
                            </tr>`;
                    });
                    
                    html += '</tbody></table></div>';
                }
                html += '</div></div>';
                document.querySelector('.rail_result').innerHTML = html;
                
                // Добавляем обработчик для кнопки экспорта в таблице
                const exportTableBtn = document.getElementById('export_rail_table');
                if (exportTableBtn) {
                    exportTableBtn.addEventListener('click', () => {
                        exportToExcel('rail');
                    });
                }
                
                hideLoading();
            })
            .catch(err => {
                console.error(err);
                hideLoading();
                alert('Ошибка при расчете');
            });
        });

        function updateSeaButton() {
            const seaPol = document.getElementById('sea_pol').value;
            const seaPod = document.getElementById('sea_pod').value;
            const seaCoc = document.getElementById('sea_coc').value;
            const seaCalculateBtn = document.getElementById('sea_calculate');
            
            if (seaPol && seaPod && seaCoc) {
                seaCalculateBtn.classList.remove('disabled');
            } else {
                seaCalculateBtn.classList.add('disabled');
            }
        }

        function updateRailButton() {
            const railOrigin = document.getElementById('rail_origin').value;
            const railDest = document.getElementById('rail_destination').value;
            const railCoc = document.getElementById('rail_coc').value;
            const railCalculateBtn = document.getElementById('rail_calculate');
            
            if (railOrigin && railDest && railCoc) {
                railCalculateBtn.classList.remove('disabled');
            } else {
                railCalculateBtn.classList.add('disabled');
            }
        }

        // Добавьте обработчики изменения полей для обновления состояния кнопок
        document.getElementById('sea_pol').addEventListener('change', updateSeaButton);
        document.getElementById('sea_pod').addEventListener('change', updateSeaButton);
        document.getElementById('sea_coc').addEventListener('change', updateSeaButton);
        document.getElementById('rail_origin').addEventListener('change', updateRailButton);
        document.getElementById('rail_destination').addEventListener('change', updateRailButton);
        document.getElementById('rail_coc').addEventListener('change', updateRailButton);
        // ЖД ПЕРЕВОЗКИ
        // Преобразуем PHP-массив zhdPerevozki в JS
        const zhdPerevozki = <?php echo json_encode($zhdPerevozki, JSON_UNESCAPED_UNICODE); ?>;

        const railOrigin      = document.getElementById('rail_origin');
        const railDestination = document.getElementById('rail_destination');
        const railCoc         = document.getElementById('rail_coc');
        const railAgent       = document.getElementById('rail_agent');
        const railHazard      = document.getElementById('rail_hazard');
        const railSecurity    = document.getElementById('rail_security');
        const railSum         = document.getElementById('rail_sum');

        function resetSelect(select, placeholder) {
            select.innerHTML = `<option value="">${placeholder}</option>`;
            select.disabled = true;
        }

        // Станция отправления → список станций назначения
        railOrigin.addEventListener('change', () => {
            resetSelect(railDestination, 'Выберите пункт...');
            resetSelect(railCoc, 'Выберите тип контейнера...');
            railSum.value = '';
            const origin = railOrigin.value;
            if (!origin) return;

            const dests = [...new Set(
                zhdPerevozki
                .filter(r => r.POL === origin)
                .map(r => r.POD)
            )];
            dests.forEach(dest => {
                const opt = document.createElement('option');
                opt.value = dest;
                opt.text  = dest;
                railDestination.appendChild(opt);
            });
            railDestination.disabled = false;
        });

        // Станция назначения → разблокировать выбор COC
        railDestination.addEventListener('change', () => {
            resetSelect(railCoc, 'Выберите тип контейнера...');
            railAgent.value = railSum.value = '';
            if (!railDestination.value) return;
            // фиксированные типы
            ['20DC (<24t)', '20DC (24t-28t)', '40HC (28t)'].forEach(type => {
                const opt = document.createElement('option');
                opt.value = type;
                opt.text  = type;
                railCoc.appendChild(opt);
            });
            railCoc.disabled = false;

            const rec = zhdPerevozki.find(r =>
                r.POL === origin && r.POD === dest
            );
            if (!rec) return;
            railAgent.value = rec.AGENT ?? '';
        });

        // На любые изменения — пересчитать
        [railCoc, railHazard, railSecurity].forEach(el => {
            el.addEventListener('change', updateRailSum);
        });

        function updateRailSum() {
            railAgent.value = railSum.value = '';
            const origin = railOrigin.value;
            const dest   = railDestination.value;
            const coc    = railCoc.value;
            const hazard = railHazard.value === 'yes';
            const secOpt = railSecurity.value; // 'no', '20' или '40'

            if (!(origin && dest && coc)) return;

            // находим запись
            const rec = zhdPerevozki.find(r =>
                r.POL === origin && r.POD === dest
            );
            if (!rec) return;

            railAgent.value = rec.AGENT ?? '';

            // выбираем базовую ставку по coc и опасности
            let cost;
            if (coc.startsWith('20DC (<24t)')) {
                cost = parseFloat(hazard ? rec.OPASNYY_20DC_24 : rec.DC20_24);
            } else if (coc.startsWith('20DC (24t-28t)')) {
                cost = parseFloat(hazard ? rec.OPASNYY_DC20_24T_28T : rec.DC20_24T_28T);
            } else if (coc.startsWith('40HC')) {
                cost = parseFloat(hazard ? rec.OPASNYY_HC40_28T : rec.HC40_28T);
            } else {
                return;
            }

            // добавляем охрану, если выбрано
            let securityCost = 0;
            if (secOpt === '20') {
                securityCost = parseFloat(rec.OKHRANA_20_FUT) || 0;
            } else if (secOpt === '40') {
                securityCost = parseFloat(rec.OKHRANA_40_FUT) || 0;
            }

            const total = cost + securityCost;
            railSum.value = total;
        }

        // КОМБИНИРОВАННЫЕ ПЕРЕВОЗКИ
        // преобразуем PHP-массив combPerevozki в JS
        const combPerevozki = <?php echo json_encode($combPerevozki, JSON_UNESCAPED_UNICODE); ?>;
        console.log(combPerevozki);
        // элементы формы
        const combPol      = document.getElementById('comb_sea_pol');
        const combDropOff      = document.getElementById('comb_drop_off');
        const combDest     = document.getElementById('comb_rail_dest');
        const combCoc      = document.getElementById('comb_coc');
        const combRemark   = document.getElementById('comb_remark');
        const combSecurity = document.getElementById('comb_security');
        const combBtn      = document.getElementById('comb_calculate');
        const combResult   = document.querySelector('.comb_result');
        const combTransshipmentPort = document.getElementById('comb_transshipment_port');

        // сброс селекта с плейсхолдером
        function resetSelect(el, placeholder) {
            el.innerHTML = `<option value="">${placeholder}</option>`;
            el.disabled = true;
        }

        // включить/выключить кнопку «Рассчитать»
        function updateCombButton() {
            const pol = combPolSelect.value;
            const dropOff = dropSelect.value;
            const dest = combDest.value;
            const coc = combCoc.value;
            
            // Порт перевалки теперь необязателен, поэтому не проверяем его
            if (pol && dropOff && dest && coc) {
                combBtn.classList.remove('disabled');
            } else {
                combBtn.classList.add('disabled');
            }
        }
        // При изменении POL, DROP OFF или станции назначения - заполняем порты перевалки
        function updateTransshipmentPorts() {
            const selectedPol = combPolSelect.value;
            const selectedDropOff = dropSelect.value;
            const selectedDest = combDest.value;
            
            // Очищаем список портов перевалки
            combTransshipmentPort.innerHTML = '<option value="">Все порты перевалки</option>';
            combTransshipmentPort.disabled = false; // Всегда доступен
            
            if (!selectedPol || !selectedDropOff || !selectedDest) {
                combBtn.classList.add('disabled');
                combTransshipmentPort.disabled = true;
                return;
            }
            
            // Проверяем, что выбранная станция назначения доступна для этого DROP OFF
            const filteredDestinations = filterDestinationsByDropOff(selectedPol, selectedDropOff);
            
            if (!filteredDestinations.includes(selectedDest)) {
                combBtn.classList.add('disabled');
                combTransshipmentPort.innerHTML = '<option value="">Станция не доступна для выбранного DROP OFF</option>';
                combTransshipmentPort.disabled = true;
                return;
            }
            
            // Получаем возможные порты перевалки для этого маршрута
            const transshipmentPorts = getTransshipmentPorts(selectedPol, selectedDropOff, selectedDest);
            
            if (transshipmentPorts.length === 0) {
                combTransshipmentPort.disabled = true;
                combTransshipmentPort.innerHTML = '<option value="">Нет доступных портов перевалки</option>';
                combBtn.classList.add('disabled');
                return;
            }
            
            // Наполняем список портов перевалки
            transshipmentPorts.forEach(port => {
                const opt = document.createElement('option');
                opt.value = port;
                opt.textContent = port;
                combTransshipmentPort.appendChild(opt);
            });

            updateCombButton();
        }

        // Функция для получения портов перевалки
        function getTransshipmentPorts(pol, dropOff, dest) {
            // 1. Находим морские записи по POL и DROP OFF
            const seaMatches = seaPerevozki.filter(item => 
                item.POL === pol && 
                item.DROP_OFF_LOCATION === dropOff
            );
            console.log(pol);
            console.log(dropOff);
            console.log(seaMatches);
            if (seaMatches.length === 0) return [];
            
            // 2. Получаем уникальные порты назначения (POD) из морских перевозок
            const seaPods = [...new Set(seaMatches.map(item => item.POD))];
            console.log(seaPods);
            /*
            // 3. Находим комбинированные записи, которые связаны с этими POD и станцией назначения
            const combMatches = combPerevozki.filter(item => 
                seaPods.includes(item.PUNKT_OTPRAVLENIYA) &&
                item.STANTSIYA_NAZNACHENIYA === dest
            );
            */
            // 4. Получаем уникальные порты отправления из комбинированных записей
            // Это и будут порты перевалки
            //const transshipmentPorts = [...new Set(seaPods.map(item => item.PUNKT_OTPRAVLENIYA))];
            
            return seaPods;
        }

        combPolSelect.addEventListener('change', () => {
            updateTransshipmentPorts();
            updateCombButton();
        });

        dropSelect.addEventListener('change', () => {
            updateTransshipmentPorts();
            updateCombButton();
        });

        combDest.addEventListener('change', () => {
            updateTransshipmentPorts();
            updateCombButton();
        });

        combTransshipmentPort.addEventListener('change', function() {
            // Находим комбинированную запись по выбранному порту перевалки (PUNKT_OTPRAVLENIYA)
            const selectedTransshipment = this.value;
            const selectedDest = combDest.value;
            
            if (selectedTransshipment && selectedDest) {
                const combMatch = combPerevozki.find(item => 
                    item.PUNKT_OTPRAVLENIYA === selectedTransshipment &&
                    item.STANTSIYA_NAZNACHENIYA === selectedDest
                );
                
                if (combMatch && combMatch.REMARK) {
                    combRemark.value = combMatch.REMARK;
                } else {
                    combRemark.value = '';
                }
            }
            
            updateCombButton();
        });

        // функция для фильтрации станций назначения по DROP OFF
        function filterDestinationsByDropOff(pol, dropOff) {
            // 1. Находим морские записи по выбранному POL и DROP OFF
            const seaMatches = seaPerevozki.filter(item => 
                item.POL === pol && 
                item.DROP_OFF_LOCATION === dropOff
            );
            
            if (seaMatches.length === 0) return [];
            const mapping = getLocationMapping();
            const mappedDropOff = mapping[dropOff] || dropOff;
            // 2. Получаем уникальные порты назначения (POD) из морских перевозок для этого DROP OFF
            const seaPods = [...new Set(seaMatches.map(item => item.POD))];
            
            // 3. Находим комбинированные записи, которые связаны с этими POD
            const filteredCombMatches = combPerevozki.filter(item => {
                // Проверяем по PUNKT_NAZNACHENIYA с учетом сопоставления
                return item.PUNKT_NAZNACHENIYA === mappedDropOff || 
                    item.PUNKT_NAZNACHENIYA === dropOff; // также проверяем оригинальное значение
            });
            
            // 4. Получаем уникальные станции назначения
            const dests = [...new Set(filteredCombMatches.map(item => item.STANTSIYA_NAZNACHENIYA))];
            
            console.log('Filtered destinations for', pol, dropOff, ':', dests);
            return dests;
        }

        // при выборе POL — наполняем станции назначения
        combPolSelect.addEventListener('change', function () {
            const selectedPol = this.value;

            // очищаем списки
            dropSelect.innerHTML = '<option value="">Выберите...</option>';
            combDest.innerHTML = '<option value="">Выберите станцию...</option>';
            resetSelect(combCoc, 'Выберите тип контейнера...');
            
            if (!selectedPol) {
                dropSelect.disabled = true;
                combDest.disabled = true;
                combBtn.classList.add('disabled');
                return;
            }

            // выбираем все элементы с этим POL
            const matches = seaPerevozki.filter(item => item.POL === selectedPol);
            
            // получаем уникальные DROP_OFF_LOCATION
            const dropOffs = [...new Set(matches.map(item => item.DROP_OFF_LOCATION))];

            // наполняем select DROP OFF
            dropOffs.forEach(loc => {
                const opt = document.createElement('option');
                opt.value = loc;
                opt.textContent = loc;
                dropSelect.appendChild(opt);
            });

            dropSelect.disabled = dropOffs.length === 0;
            
            // Наполняем список станций назначения из комбинированных перевозок
            // используем все записи, где POL совпадает или PUNKT_OTPRAVLENIYA совпадает
            const allCombMatches = combPerevozki.filter(item => 
                item.POL === selectedPol || item.PUNKT_OTPRAVLENIYA === selectedPol
            );
            
            const dests = [...new Set(allCombMatches.map(item => item.STANTSIYA_NAZNACHENIYA))];
            
            dests.forEach(dest => {
                const opt = document.createElement('option');
                opt.value = dest;
                opt.textContent = dest;
                combDest.appendChild(opt);
            });
            
            combDest.disabled = dests.length === 0;
        });

                // при выборе DROP OFF — наполняем станции назначения
        dropSelect.addEventListener('change', function() {
            const selectedPol = combPolSelect.value;
            const selectedDropOff = this.value;
            
            // очищаем списки
            resetSelect(combDest, 'Выберите станцию назначения...');
            resetSelect(combCoc, 'Выберите тип контейнера...');
            combRemark.value = '';
            combTransshipmentPort.innerHTML = '<option value="">Выберите порт перевалки...</option>';
            combTransshipmentPort.disabled = true;
            combBtn.classList.add('disabled');
            
            if (!selectedPol || !selectedDropOff) {
                combDest.disabled = true;
                return;
            }
            
            // Получаем отфильтрованные станции назначения по DROP OFF
            const filteredDestinations = filterDestinationsByDropOff(selectedPol, selectedDropOff);
            
            if (filteredDestinations.length === 0) {
                combDest.disabled = true;
                combDest.innerHTML = '<option value="">Нет доступных станций для выбранного DROP OFF</option>';
                return;
            }
            
            // Наполняем список станций назначения
            filteredDestinations.forEach(dest => {
                const opt = document.createElement('option');
                opt.value = dest;
                opt.textContent = dest;
                combDest.appendChild(opt);
            });
            
            combDest.disabled = false;
        });
        // при выборе станции — активируем выбор COC
        combDest.addEventListener('change', function() {
            const selectedPol = combPolSelect.value;
            const selectedDropOff = dropSelect.value;
            const selectedDest = this.value;
            
            resetSelect(combCoc, 'Выберите тип контейнера...');
            combRemark.value = '';
            combBtn.classList.add('disabled');
            
            if (!selectedDest || !selectedPol || !selectedDropOff) {
                combCoc.disabled = true;
                return;
            }
            
            // Наполняем список типов контейнеров
            ['20DC (<24t)', '20DC (24t-28t)', '40HC (28t)'].forEach(type => {
                const opt = document.createElement('option');
                opt.value = type;
                opt.textContent = type;
                combCoc.appendChild(opt);
            });
            
            combCoc.disabled = false;
            
            // Если нужно, можно заполнить Remark из комбинированных записей
            const filteredDestinations = filterDestinationsByDropOff(selectedPol, selectedDropOff);
            if (filteredDestinations.includes(selectedDest)) {
                // Находим подходящую комбинированную запись
                const seaMatches = seaPerevozki.filter(item => 
                    item.POL === selectedPol && 
                    item.DROP_OFF_LOCATION === selectedDropOff
                );
                
                const seaPods = [...new Set(seaMatches.map(item => item.POD))];
                
                const combMatch = combPerevozki.find(item => 
                    (seaPods.includes(item.POL) || seaPods.includes(item.PUNKT_OTPRAVLENIYA)) &&
                    item.STANTSIYA_NAZNACHENIYA === selectedDest
                );
                
                if (combMatch && combMatch.REMARK) {
                    combRemark.value = combMatch.REMARK;
                }
            }
            
            updateCombButton();
        });

        combCoc.addEventListener('change', () => {
            // сначала очищаем Remark
            combRemark.value = '';

            // если контейнер выбран и ранее выбраны POL + Dest
            if (combPol.value && combDest.value && combCoc.value) {
                // ищем первую подходящую запись
                const match = combPerevozki.find(item =>
                    // пока коментим сложную логику
                    //item.POL === combPol.value &&
                    item.STANTSIYA_NAZNACHENIYA === combDest.value 
                );

                // если нашли и есть Remark
                if (match && match.REMARK) {
                    combRemark.value = match.REMARK;
                }
            }

            // после изменения COC или Remark/Security — проверяем, активна ли кнопка
            updateCombButton();
        });

        // на любые изменения COC/Remark/Security — проверяем, можно ли рассчитывать
        [combCoc, combRemark, combSecurity].forEach(el =>
            el.addEventListener('change', updateCombButton)
        );

        // при клике на «Рассчитать» — AJAX-запрос
        combBtn.addEventListener('click', e => {
            if (combBtn.classList.contains('disabled')) return;
            showLoading('Расчет комбинированного маршрута...');
            // собираем данные
            const payload = {
                comb_sea_pol: document.getElementById('comb_sea_pol').value,
                comb_drop_off: document.getElementById('comb_drop_off').value,
                comb_rail_dest: document.getElementById('comb_rail_dest').value,
                comb_coc: document.getElementById('comb_coc').value,
                comb_container_ownership: document.getElementById('comb_container_ownership').value,
                comb_hazard: document.getElementById('comb_hazard').value,
                comb_transshipment_port: document.getElementById('comb_transshipment_port').value || '',
                comb_security: document.getElementById('comb_security').value,
                sea_profit: document.getElementById('sea_profit').value,
                rail_profit: document.getElementById('rail_profit').value
            };
            
            // отправляем
            fetch(cleanUrl + '?action=getCombPerevozki', {
                method: 'POST',
                headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                body: new URLSearchParams(payload)
            })
            .then(res => res.json())
            .then(data => {
                console.log(data);
                
                // Сохраняем результаты для экспорта
                currentCombResults = data;
                
                // Показываем кнопку экспорта
                const exportBtn = document.getElementById('export_comb');
                if (exportBtn) {
                    exportBtn.style.display = 'inline-block';
                }
                
                const combResult = document.querySelector('.comb_result');
                combResult.innerHTML = ''; // очистим предыдущий результат

                if (!Array.isArray(data) || data.length === 0) {
                    combResult.innerHTML = '<div class="alert alert-warning">Нет доступных маршрутов</div>';
                    return;
                }
                
                let html = `
            <div class="card mt-4">
                <div class="card-header d-flex justify-content-between align-items-center">
                    <h4 class="mb-0">Результаты расчета комбинированных перевозок</h4>
                    <div>
                        <button type="button" id="export_comb_table" class="btn btn-success btn-sm export-btn">
                            <i class="bi bi-file-excel"></i> Экспорт в Excel
                        </button>
                    </div>
                </div>
                <div class="card-body">
                    <div class="table-responsive">
                        <table class="table table-bordered table-striped">
                            <thead class="table-light">
                                <tr>
                                    <th>Морской порт отправления</th>
                                    <th>Морской порт прибытия</th>
                                    <th>ЖД станция отправления</th>
                                    <th>ЖД станция назначения</th>
                                    <th>DROP OFF LOCATION</th>
                                    <th>Тип контейнера</th>
                                    <th>Собственность контейнера</th>
                                    <th>Опасный груз</th>
                                    <th>Охрана</th>
                                    <th>Стоимость морской части ($)</th>
                                    <th>Стоимость ЖД части (₽)</th>
                                    <th>Стоимость ЖД части опасный (₽)</th>
                                    <th>Общая стоимость ($ + ₽)</th>
                                    <th>Общая стоимость опасный ($ + ₽)</th>
                                    <th>Агент(-ы)</th>
                                    <th>Комментарий</th>
                                </tr>
                            </thead>
                            <tbody>`;
                data.forEach(item => {
                    html += `
                        <tr>
                            <td>${item.comb_sea_pol || ''}</td>
                            <td>${item.comb_sea_pod || ''}</td>
                            <td>${item.comb_rail_start || ''}</td>
                            <td>${item.comb_rail_dest || ''}</td>
                            <td>${item.drop_off_location || ''}</td>
                            <td>${item.comb_coc || ''}</td>
                            <td>${item.comb_container_ownership || ''}</td>
                            <td>${item.comb_hazard || 'Нет'}</td>
                            <td>${item.comb_security || ''}</td>
                            <td><strong>${item.cost_sea || 0}</strong></td>
                            <td><strong>${item.cost_rail || 0}</strong></td>
                            <td><strong>${item.cost_rail_danger || 0}</strong></td>
                            <td><strong>${item.cost_total || 0}</strong></td>
                            <td><strong>${item.cost_total_danger || item.cost_total || 0}</strong></td>
                            <td>${item.agent || ''}</td>
                            <td>${item.remark || ''}</td>
                        </tr>`;
                });
                html += '</tbody></table></div></div></div>';
                combResult.innerHTML = html;
                
                // Добавляем обработчик для кнопки экспорта в таблице
                const exportTableBtn = document.getElementById('export_comb_table');
                if (exportTableBtn) {
                    exportTableBtn.addEventListener('click', () => {
                        exportToExcel('combined');
                    });
                }
                
                hideLoading();
            })
            .catch(err => {
                console.error(err);
                hideLoading();
                combResult.innerHTML = '<div class="alert alert-danger">Ошибка при запросе.</div>';
            });
        });

        // ===== ЗАГРУЗКА ФАЙЛОВ =====
        // соответствие id кнопок <→> action-параметру
        const uploadMap = {
        'upload_zhd':  'uploadZhd',
        'upload_sea':  'uploadSea',
        'upload_comb': 'uploadComb'
        };

        const fileInput = document.getElementById('upload_file_input');
        fileInput.setAttribute('accept', '.xlsx,.xls'); // ограничим диалог
        let currentAction = null;

        const uploadButtons = Object.keys(uploadMap).map(id => document.getElementById(id));
        const setUploading = (on) => uploadButtons.forEach(b => b && (b.disabled = on));

        const showMessage = (text, type = 'info') => {
        // Если есть элемент для статуса — используем его, иначе alert
        const box = document.getElementById('upload_status');
        if (box) {
            box.textContent = text;
            box.dataset.type = type; // можно стилизовать по [data-type]
        } else {
            if (type === 'error') alert(text); else console.log(text);
        }
        };

        // безопасный разбор JSON даже при ошибках
        const parseResponse = async (response) => {
        const raw = await response.text();
        let payload;
        try {
            payload = raw ? JSON.parse(raw) : {};
        } catch {
            payload = { error: raw || response.statusText };
        }
        return { ok: response.ok, status: response.status, payload };
        };

        // таймаут запроса
        const withTimeout = (ms) => {
        const controller = new AbortController();
        const t = setTimeout(() => controller.abort(), ms);
        return { controller, clear: () => clearTimeout(t) };
        };

        // Навешиваем на каждую кнопку открытие file-dialog
        Object.keys(uploadMap).forEach(btnId => {
        const btn = document.getElementById(btnId);
        btn.addEventListener('click', () => {
            currentAction = uploadMap[btnId];
            fileInput.value = ''; // сброс предыдущего выбора
            fileInput.click();    // открыть диалог
        });
        });
        
        // Функция экспорта в Excel
        function exportToExcel(type) {
            let data = [];
            let action = '';
            
            switch(type) {
                case 'sea':
                    data = currentSeaResults;
                    action = 'exportSeaToExcel';
                    break;
                case 'rail':
                    data = currentRailResults;
                    action = 'exportRailToExcel';
                    break;
                case 'combined':
                    data = currentCombResults;
                    action = 'exportCombToExcel';
                    break;
                default:
                    return;
            }
            
            if (!data || data.length === 0) {
                alert('Нет данных для экспорта');
                return;
            }
            
            showLoading('Подготовка файла Excel...');
            
            fetch(cleanUrl + '?action=' + action, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify(data)
            })
            .then(response => {
                if (response.ok) {
                    return response.blob();
                } else {
                    return response.json().then(err => { throw new Error(err.message || 'Ошибка экспорта'); });
                }
            })
            .then(blob => {
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.style.display = 'none';
                a.href = url;
                
                // Генерируем имя файла
                const date = new Date().toISOString().slice(0, 19).replace(/:/g, '-');
                let filename = '';
                switch(type) {
                    case 'sea':
                        filename = `морские_перевозки_${date}.xlsx`;
                        break;
                    case 'rail':
                        filename = `жд_перевозки_${date}.xlsx`;
                        break;
                    case 'combined':
                        filename = `комбинированные_перевозки_${date}.xlsx`;
                        break;
                }
                a.download = filename;
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                document.body.removeChild(a);
                hideLoading();
                
                // Показываем сообщение об успешном экспорте
                showUploadMessage(`Файл "${filename}" успешно экспортирован`, 'success');
            })
            .catch(error => {
                console.error('Ошибка экспорта:', error);
                hideLoading();
                alert('Ошибка при экспорте в Excel: ' + error.message);
            });
        }
        const seaExportButton = document.getElementById('export_sea');
        const railExportButton = document.getElementById('export_rail');
        const combExportButton = document.getElementById('export_comb');

        if (seaExportButton) {
            seaExportButton.addEventListener('click', () => exportToExcel('sea'));
        }

        if (railExportButton) {
            railExportButton.addEventListener('click', () => exportToExcel('rail'));
        }

        if (combExportButton) {
            combExportButton.addEventListener('click', () => exportToExcel('combined'));
        }
        // Функции для управления индикатором загрузки
        function showLoading(message = 'Идет расчет...') {
            const overlay = document.getElementById('loading-overlay');
            const messageEl = document.getElementById('loading-message');
            
            if (overlay && messageEl) {
                messageEl.textContent = message;
                overlay.style.display = 'block';
            }
        }

        function hideLoading() {
            const overlay = document.getElementById('loading-overlay');
            if (overlay) {
                overlay.style.display = 'none';
            }
        }

        function updateProgress(percent, message = '') {
            const progressBar = document.getElementById('loading-progress');
            const messageEl = document.getElementById('loading-message');
            
            if (progressBar) {
                progressBar.style.width = Math.min(100, Math.max(0, percent)) + '%';
            }
            
            if (message && messageEl) {
                messageEl.textContent = message;
            }
        }

        // Функция для показа сообщений о загрузке файлов
        function showUploadMessage(text, type = 'info') {
            const statusBox = document.getElementById('upload-status');
            const statusText = document.getElementById('upload-status-text');
            
            if (statusBox && statusText) {
                // Устанавливаем текст и классы стилей
                statusText.textContent = text;
                
                // Убираем все классы alert-* и добавляем нужный
                statusBox.className = 'alert alert-dismissible fade show mb-4';
                switch(type) {
                    case 'success':
                        statusBox.classList.add('alert-success');
                        break;
                    case 'error':
                        statusBox.classList.add('alert-danger');
                        break;
                    case 'warning':
                        statusBox.classList.add('alert-warning');
                        break;
                    default:
                        statusBox.classList.add('alert-info');
                }
                
                // Показываем блок
                statusBox.style.display = 'block';
                
                // Автоматически скрываем через 5 секунд для success/info
                if (type !== 'error') {
                    setTimeout(() => {
                        hideUploadMessage();
                    }, 5000);
                }
            }
        }

        function hideUploadMessage() {
            const statusBox = document.getElementById('upload-status');
            if (statusBox) {
                statusBox.style.display = 'none';
            }
        }

        // Закрытие по клику на крестик
        document.addEventListener('click', function(e) {
            if (e.target.classList.contains('btn-close')) {
                const alert = e.target.closest('.alert');
                if (alert) {
                    alert.style.display = 'none';
                }
            }
        });
        // Когда файл выбран — отправляем
        fileInput.addEventListener('change', async () => {
        if (!fileInput.files.length || !currentAction) return;

        const file = fileInput.files[0];
        // Определяем тип операции для сообщения
        const operationNames = {
            'uploadZhd': 'Ж/Д перевозок',
            'uploadSea': 'морских перевозок',
            'uploadComb': 'комбинированных перевозок'
        };
        
        const operationName = operationNames[currentAction] || 'данных';
        // Показываем индикатор загрузки
        showLoading(`Загрузка файла ${file.name}...`);
        updateProgress(10, `Начало загрузки файла`);
        // Небольшая предвалидация на клиенте
        const maxSizeMb = 32; // синхронизируй с upload_max_filesize/post_max_size
        if (file.size > maxSizeMb * 1024 * 1024) {
            showMessage(`Файл больше ${maxSizeMb} МБ. Уменьшите размер или увеличьте лимиты на сервере.`, 'error');
            return;
        }

        const formData = new FormData();
        formData.append('file', file);

        const { controller, clear } = withTimeout(120000); // 120 сек таймаут

        setUploading(true);
        updateProgress(30, 'Отправка файла на сервер...');
        showMessage(`Загружаю файл (${file.name}) для действия ${currentAction}…`, 'info');

        try {
            const res = await fetch(`${cleanUrl}?action=${encodeURIComponent(currentAction)}`, {
            method: 'POST',
            body: formData,
            signal: controller.signal,
            // если нужен cookie для авторизации и у тебя subdomain тот же:
            credentials: 'same-origin'
            });
            updateProgress(70, 'Обработка файла на сервере...');
            const { ok, status, payload } = await parseResponse(res);
            hideLoading();
            // Успехи/частичные успехи
            if (ok && (status === 200 || status === 207)) {
            // ожидаемый payload со стороны бэка:
            // { result: boolean, added: number, errors: [{row, error}]?, message: string }
            const added = payload.added ?? 0;
            const errors = Array.isArray(payload.errors) ? payload.errors : [];
            if (status === 207 || errors.length) {
                const firstErrors = errors.slice(0, 5).map(e => `строка ${e.row}: ${e.error}`).join('\n');
                const more = errors.length > 5 ? `\n…и ещё ${errors.length - 5}` : '';
                showMessage(
                `Загрузка завершена с частичными ошибками.\nДобавлено: ${added}.\nОшибки:\n${firstErrors}${more}`,
                'error'
                );
            } else {
                showMessage(payload.message || `Готово! Добавлено записей: ${added}.`, 'info');
            }
            console.log(`Результат ${currentAction}:`, payload);
            return;
            }

            // Обработка типичных кодов ошибок
            const errText = payload?.error || payload?.message || res.statusText || 'Неизвестная ошибка';
            switch (status) {
            case 400:
                showMessage(`Некорректный запрос: ${errText}`, 'error'); break;
            case 405:
                showMessage(`Метод не разрешён: ${errText}`, 'error'); break;
            case 413:
                showMessage(`Файл слишком большой (413): ${errText}. Уменьшите файл или поднимите лимиты на сервере.`, 'error'); break;
            case 415:
                showMessage(`Неподдерживаемый тип файла: ${errText}`, 'error'); break;
            case 422:
                showMessage(`Файл распознан, но данные некорректны: ${errText}`, 'error'); break;
            case 500:
                showMessage(`Серверная ошибка (500): ${errText}`, 'error'); break;
            default:
                showMessage(`Ошибка ${status}: ${errText}`, 'error');
            }
        } catch (e) {
            hideLoading();
            if (e.name === 'AbortError') {
            showMessage('Превышено время ожидания запроса. Попробуйте снова или уменьшите файл.', 'error');
            } else {
            showMessage(`Сетевая ошибка: ${e.message || e}`, 'error');
            }
        } finally {
            clear();
            setUploading(false);
            currentAction = null;
        }
        });
    });
</script>
</body>
</html>