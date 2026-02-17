<?php
/**
 * Форма калькулятора расчета стоимости перевозок (Vuetify) с зависимым списком станций Ж/Д перевозок
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
$allSeaPods = array_unique(array_column($seaPerevozki, 'POD'));

// Подготовка данных для Vue
$seaPortsForVue = array_values($seaPorts);
$zhdStartsForVue = array_values($zhdStarts);
$combStartsForVue = array_values($combStarts);
?>
<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Калькулятор стоимости перевозок</title>
    
    <!-- Vuetify CSS -->
    <link href="https://cdn.jsdelivr.net/npm/vuetify@3.5.0/dist/vuetify.min.css" rel="stylesheet">
    <link href="https://cdn.jsdelivr.net/npm/@mdi/font@7.4.47/css/materialdesignicons.min.css" rel="stylesheet">
    <link href="https://fonts.googleapis.com/css?family=Roboto:100,300,400,500,700,900&display=swap" rel="stylesheet">
    
    <script src="https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/file-saver@2.0.5/dist/FileSaver.min.js"></script>
    <style>
        .v-application {
            background-color: #f5f5f5 !important;
        }
        .loading-overlay {
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(0,0,0,0.5);
            z-index: 9999;
            display: flex;
            justify-content: center;
            align-items: center;
        }
        .result-table {
            margin-top: 20px;
        }
        .section-card {
            margin-bottom: 20px;
        }
    </style>
</head>
<body>
<div id="app">
    <v-app>
        <!-- Индикатор загрузки -->
        <div v-if="loading" class="loading-overlay">
            <v-card class="pa-4 text-center" width="400">
                <v-progress-circular
                    indeterminate
                    color="primary"
                    size="64"
                    class="mb-3"
                ></v-progress-circular>
                <h5 class="mb-2">{{ loadingMessage }}</h5>
                <v-progress-linear
                    v-if="loadingProgress > 0"
                    :model-value="loadingProgress"
                    color="primary"
                    height="6"
                    class="mt-3"
                ></v-progress-linear>
            </v-card>
        </div>

        <!-- Сообщения о загрузке файлов -->
        <!--
        <v-alert
            v-if="uploadMessage"
            :type="uploadMessageType"
            class="upload-status mb-4 mx-4 mt-4"
            dismissible
            @click:close="clearUploadMessage"
        >
            {{ uploadMessage }}
        </v-alert>
 -->
        <v-main>
            <v-container fluid class="pa-4">
                <v-card class="pa-4">
                    <!-- Заголовок и кнопки загрузки -->
                    <v-row class="mb-4" align="center">
                        <v-col cols="12" md="6">
                            <h1 class="text-h4 font-weight-bold">Калькулятор стоимости перевозок</h1>
                        </v-col>
                        <v-col cols="12" md="6" class="text-md-right">
                            <v-btn
                                color="primary"
                                class="mr-2 mb-2"
                                @click="openFileDialog('comb')"
                                :loading="uploading"
                            >
                                <v-icon start>mdi-upload</v-icon>
                                Комбинированный
                            </v-btn>
                            <v-btn
                                color="primary"
                                class="mr-2 mb-2"
                                @click="openFileDialog('sea')"
                                :loading="uploading"
                            >
                                <v-icon start>mdi-upload</v-icon>
                                Морские
                            </v-btn>
                            <v-btn
                                color="primary"
                                class="mb-2"
                                @click="openFileDialog('zhd')"
                                :loading="uploading"
                            >
                                <v-icon start>mdi-upload</v-icon>
                                Ж/Д
                            </v-btn>
                        </v-col>
                    </v-row>

                    <!-- Выбор типа расчета -->
                    <v-row class="mb-6">
                        <v-col cols="12">
                            <v-select
                                v-model="calcType"
                                label="Выберите тип расчёта"
                                :items="calcTypes"
                                variant="outlined"
                                required
                                @update:model-value="onCalcTypeChange"
                            ></v-select>
                        </v-col>
                    </v-row>

                    <!-- МОРСКИЕ ПЕРЕВОЗКИ -->
                    <v-card v-if="calcType === 'sea'" class="section-card">
                        <v-card-title class="d-flex justify-space-between align-center">
                            <span>Морские перевозки</span>
                        </v-card-title>
                        <v-card-text>
                            <v-row>
                                <v-col cols="12" md="6">
                                    <v-select
                                        v-model="seaForm.pol"
                                        label="POL (Порт отправления)"
                                        :items="seaPorts"
                                        variant="outlined"
                                        @update:model-value="onSeaPolChange"
                                        item-title="title"
                                        item-value="value"
                                        :return-object="false"
                                    ></v-select>
                                </v-col>
                                <v-col cols="12" md="6">
                                    <v-select
                                        v-model="seaForm.pod"
                                        label="POD (Порт прибытия)"
                                        :items="seaPods"
                                        variant="outlined"
                                        :disabled="seaFormDisabled.pod"
                                        @update:model-value="onSeaPodChange"
                                        item-title="title"
                                        item-value="value"
                                        :return-object="false"
                                    ></v-select>
                                </v-col>
                                <v-col cols="12" md="3">
                                    <v-select
                                        v-model="seaForm.coc"
                                        label="ТИП КОНТЕЙНЕРА"
                                        :items="seaCocTypes"
                                        variant="outlined"
                                        :disabled="seaFormDisabled.coc"
                                        @update:model-value="onSeaCocChange"
                                        item-title="title"
                                        item-value="value"
                                        :return-object="false"
                                    ></v-select>
                                </v-col>
                                <v-col cols="12" md="6">
                                    <v-select
                                        v-model="seaForm.dropOffLocation"
                                        label="DROP OFF LOCATION"
                                        :items="seaDropOffLocations"
                                        variant="outlined"
                                        :disabled="seaFormDisabled.dropOffLocation"
                                        @update:model-value="onSeaDropOffChange"
                                        item-title="title"
                                        item-value="value"
                                        :return-object="false"
                                    ></v-select>
                                </v-col>
                                <v-col cols="12" md="3">
                                    <v-select
                                        v-model="seaForm.containerOwnership"
                                        label="СОБСТВЕННОСТЬ КОНТЕЙНЕРА"
                                        :items="ownershipOptions"
                                        variant="outlined"
                                        :disabled="seaFormDisabled.containerOwnership"
                                        @update:model-value="updateSeaCosts"
                                        item-title="title"
                                        item-value="value"
                                        :return-object="false"
                                    ></v-select>
                                </v-col>
                                <v-col cols="12" md="4">
                                    <v-text-field
                                        v-model="seaForm.caf"
                                        label="% CAF (конверт)"
                                        type="number"
                                        step="0.5"
                                        :disabled="seaFormDisabled.caf"
                                        variant="outlined"
                                        @input="updateSeaCosts"
                                    ></v-text-field>
                                </v-col>
                                <v-col cols="12" md="4">
                                    <v-text-field
                                        v-model="seaForm.profit"
                                        label="Profit (Море, $)"
                                        type="number"
                                        step="1"
                                        variant="outlined"
                                        @input="updateSeaBrutto"
                                        :disabled="seaFormDisabled.profit"
                                    ></v-text-field>
                                </v-col>
                                <v-col cols="12" class="text-center">
                                    <v-btn
                                        color="primary"
                                        size="large"
                                        @click="calculateSea"
                                        :disabled="!canCalculateSea"
                                        :loading="calculatingSea"
                                    >
                                        Рассчитать
                                    </v-btn>
                                </v-col>
                            </v-row>

                            <!-- Результаты морских перевозок -->
                            <div v-if="seaResults.length > 0" class="result-table">
                                <v-card>
                                    <v-card-title class="d-flex justify-space-between align-center">
                                        <span>Результаты расчета морских перевозок</span>
                                        <v-btn
                                            color="success"
                                            @click="exportToExcel('sea')"
                                        >
                                            <v-icon start>mdi-file-excel</v-icon>
                                            Экспорт в Excel
                                        </v-btn>
                                    </v-card-title>
                                    <v-card-text>
                                        <v-data-table
                                            :headers="seaResultHeaders"
                                            :items="seaResults"
                                            class="elevation-1"
                                        ></v-data-table>
                                    </v-card-text>
                                </v-card>
                            </div>
                        </v-card-text>
                    </v-card>

                    <!-- Ж/Д ПЕРЕВОЗКИ -->
                    <v-card v-if="calcType === 'rail'" class="section-card">
                        <v-card-title class="d-flex justify-space-between align-center">
                            <span>Ж/Д перевозки</span>
                        </v-card-title>
                        <v-card-text>
                            <v-row>
                                <v-col cols="12" md="6">
                                    <v-select
                                        v-model="railForm.origin"
                                        label="Станция отправления"
                                        :items="zhdStarts"
                                        variant="outlined"
                                        @update:model-value="onRailOriginChange"
                                        item-title="title"
                                        item-value="value"
                                        :return-object="false"
                                    ></v-select>
                                </v-col>
                                <v-col cols="12" md="6">
                                    <v-select
                                        v-model="railForm.destination"
                                        label="Станция назначения"
                                        :items="railDestinations"
                                        variant="outlined"
                                        :disabled="railFormDisabled.destination"
                                        @update:model-value="onRailDestinationChange"
                                        item-title="title"
                                        item-value="value"
                                        :return-object="false"
                                    ></v-select>
                                </v-col>
                                <v-col cols="12" md="4">
                                    <v-select
                                        v-model="railForm.coc"
                                        label="ТИП КОНТЕЙНЕРА"
                                        :items="railCocTypes"
                                        variant="outlined"
                                        :disabled="railFormDisabled.coc"
                                        item-title="title"
                                        item-value="value"
                                        :return-object="false"
                                    ></v-select>
                                </v-col>
                                <v-col cols="12" md="3">
                                    <v-select
                                        v-model="railForm.containerOwnership"
                                        label="СОБСТВЕННОСТЬ КОНТЕЙНЕРА"
                                        :items="ownershipOptions"
                                        variant="outlined"
                                        item-title="title"
                                        item-value="value"
                                        :return-object="false"
                                        :disabled="railFormDisabled.containerOwnership"
                                    ></v-select>
                                </v-col>
                                <v-col cols="12" md="3">
                                    <v-select
                                        v-model="railForm.security"
                                        label="Охрана"
                                        :items="securityOptions"
                                        :disabled="railFormDisabled.security"
                                        variant="outlined"
                                        item-title="title"
                                        item-value="value"
                                        :return-object="false"
                                    ></v-select>
                                </v-col>
                                <v-col cols="12" md="4">
                                    <v-text-field
                                        v-model="railForm.profit"
                                        label="Profit (ЖД, ₽)"
                                        type="number"
                                        step="1"
                                        variant="outlined"
                                        :disabled="railFormDisabled.profit"
                                    ></v-text-field>
                                </v-col>
                                <v-col cols="12" class="text-center">
                                    <v-btn
                                        color="primary"
                                        size="large"
                                        @click="calculateRail"
                                        :disabled="!canCalculateRail"
                                        :loading="calculatingRail"
                                    >
                                        Рассчитать
                                    </v-btn>
                                </v-col>
                            </v-row>

                            <!-- Результаты Ж/Д перевозок -->
                            <div v-if="railResults.length > 0" class="result-table">
                                <v-card>
                                    <v-card-title class="d-flex justify-space-between align-center">
                                        <span>Результаты расчета Ж/Д перевозок</span>
                                        <v-btn
                                            color="success"
                                            @click="exportToExcel('rail')"
                                        >
                                            <v-icon start>mdi-file-excel</v-icon>
                                            Экспорт в Excel
                                        </v-btn>
                                    </v-card-title>
                                    <v-card-text>
                                        <v-data-table
                                            :headers="railResultHeaders"
                                            :items="railResults"
                                            class="elevation-1"
                                        ></v-data-table>
                                    </v-card-text>
                                </v-card>
                            </div>
                        </v-card-text>
                    </v-card>

                    <!-- КОМБИНИРОВАННЫЙ МАРШРУТ -->
                    <v-card v-if="calcType === 'combined'" class="section-card">
                        <v-card-title class="d-flex justify-space-between align-center">
                            <span>Комбинированный маршрут</span>
                        </v-card-title>
                        <v-card-text>
                            <v-row>
                                <v-col cols="12" md="4">
                                    <v-select
                                        v-model="combForm.seaPol"
                                        label="POL (Порт отправления)"
                                        :items="combStarts"
                                        variant="outlined"
                                        @update:model-value="onCombPolChange"
                                        item-title="title"
                                        item-value="value"
                                        :return-object="false"
                                    ></v-select>
                                </v-col>
                                <v-col cols="12" md="4">
                                    <v-select
                                        v-model="combForm.dropOff"
                                        label="DROP OFF"
                                        :items="combDropOffs"
                                        variant="outlined"
                                        :disabled="combFormDisabled.dropOff"
                                        @update:model-value="onCombDropOffChange"
                                        item-title="title"
                                        item-value="value"
                                        :return-object="false"
                                    ></v-select>
                                </v-col>
                                <v-col cols="12" md="4">
                                    <v-select
                                        v-model="combForm.destination"
                                        label="ПУНКТ НАЗНАЧЕНИЯ"
                                        :items="combDestinations"
                                        variant="outlined"
                                        :disabled="combFormDisabled.destination"
                                        item-title="title"
                                        item-value="value"
                                    ></v-select>
                                </v-col>
                                <v-col cols="12" md="5">
                                    <v-select
                                        v-model="combForm.transshipmentPort"
                                        label="ПОРТ ПЕРЕВАЛКИ"
                                        :items="transshipmentPorts"
                                        variant="outlined"
                                        :disabled="combFormDisabled.transshipmentPort"
                                        @update:model-value="onTransshipmentPortChange"
                                        item-title="title"
                                        item-value="value"
                                        :return-object="false"
                                        clearable
                                    ></v-select>
                                </v-col>
                                <v-col cols="12" md="4">
                                    <v-select
                                        v-model="combForm.coc"
                                        label="ТИП КОНТЕЙНЕРА"
                                        :items="combCocTypes"
                                        variant="outlined"
                                        :disabled="combFormDisabled.coc"
                                        item-title="title"
                                        item-value="value"
                                        :return-object="false"
                                    ></v-select>
                                </v-col>
                                <v-col cols="12" md="3">
                                    <v-select
                                        v-model="combForm.containerOwnership"
                                        label="СОБСТВЕННОСТЬ КОНТЕЙНЕРА"
                                        :items="ownershipOptions"
                                        variant="outlined"
                                        item-title="title"
                                        item-value="value"
                                        :return-object="false"
                                        :disabled="combFormDisabled.containerOwnership"
                                    ></v-select>
                                </v-col>
                                <v-col cols="12" md="4">
                                    <v-select
                                        v-model="combForm.hazard"
                                        label="Опасный груз?"
                                        :items="hazardOptions"
                                        variant="outlined"
                                        item-title="title"
                                        item-value="value"
                                        :return-object="false"
                                        :disabled="combFormDisabled.hazard"
                                    ></v-select>
                                </v-col>
                                <v-col cols="12" md="6">
                                    <v-select
                                        v-model="combForm.security"
                                        label="Охрана"
                                        :items="combSecurityOptions"
                                        variant="outlined"
                                        item-title="title"
                                        item-value="value"
                                        :return-object="false"
                                        :disabled="combFormDisabled.security"
                                    ></v-select>
                                </v-col>
                                <v-col cols="12" md="4">
                                    <v-text-field
                                        v-model="combForm.seaProfit"
                                        label="Profit (Море, $)"
                                        type="number"
                                        step="1"
                                        variant="outlined"
                                        :disabled="combFormDisabled.seaProfit"
                                    ></v-text-field>
                                </v-col>
                                <v-col cols="12" md="4">
                                    <v-text-field
                                        v-model="combForm.railProfit"
                                        label="Profit (ЖД, ₽)"
                                        type="number"
                                        step="1"
                                        variant="outlined"
                                        :disabled="combFormDisabled.railProfit"
                                    ></v-text-field>
                                </v-col>
                                <v-col cols="12" class="text-center">
                                    <v-btn
                                        color="primary"
                                        size="large"
                                        @click="calculateCombined"
                                        :disabled="!canCalculateCombined"
                                        :loading="calculatingCombined"
                                    >
                                        Рассчитать
                                    </v-btn>
                                </v-col>
                            </v-row>

                            <!-- Результаты комбинированных перевозок -->
                            <div v-if="combResults.length > 0" class="result-table">
                                <v-card>
                                    <v-card-title class="d-flex justify-space-between align-center">
                                        <span>Результаты расчета комбинированных перевозок</span>
                                        <v-btn
                                            color="success"
                                            @click="exportToExcel('combined')"
                                        >
                                            <v-icon start>mdi-file-excel</v-icon>
                                            Экспорт в Excel
                                        </v-btn>
                                    </v-card-title>
                                    <v-card-text>
                                        <v-data-table
                                            :headers="combResultHeaders"
                                            :items="combResults"
                                            class="elevation-1"
                                        ></v-data-table>
                                    </v-card-text>
                                </v-card>
                            </div>
                        </v-card-text>
                    </v-card>

                    <!-- Скрытый input для загрузки файлов -->
                    <input
                        type="file"
                        ref="fileInput"
                        accept=".xlsx,.xls"
                        style="display: none"
                        @change="onFileSelected"
                    />
                </v-card>
            </v-container>
        </v-main>
    </v-app>
</div>

<!-- Vue 3 и Vuetify -->
<script src="https://cdn.jsdelivr.net/npm/vue@3.3.4/dist/vue.global.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/vuetify@3.5.0/dist/vuetify.min.js"></script>

<script>
const { createApp, ref, computed } = Vue;
const { VDataTable } = Vuetify;

const cleanUrl = window.location.origin + window.location.pathname;

createApp({
    components: {
        VDataTable
    },
    setup() {
        // Данные из PHP
        const seaPerevozkiData = <?= json_encode($seaPerevozki, JSON_UNESCAPED_UNICODE) ?>;
        const zhdPerevozkiData = <?= json_encode($zhdPerevozki, JSON_UNESCAPED_UNICODE) ?>;
        const combPerevozkiData = <?= json_encode($combPerevozki, JSON_UNESCAPED_UNICODE) ?>;
        
        // Преобразуем массивы для удобства использования
        const seaPerevozki = Array.isArray(seaPerevozkiData) ? seaPerevozkiData : [];
        const zhdPerevozki = Array.isArray(zhdPerevozkiData) ? zhdPerevozkiData : [];
        const combPerevozki = Array.isArray(combPerevozkiData) ? combPerevozkiData : [];

        // Состояние приложения
        const calcType = ref('');
        const loading = ref(false);
        const loadingMessage = ref('Идет расчет...');
        const loadingProgress = ref(0);
        const uploading = ref(false);
        const uploadMessage = ref('');
        const uploadMessageType = ref('info');
        const combSeaPods = ref([]);
        // Состояние для форм
        const seaForm = ref({
            pol: '',
            pod: '',
            coc: '',
            hazard: 'no',
            security: 'no',
            containerOwnership: 'no',
            dropOffLocation: '',
            caf: '',
            profit: '',
            netto: '',
            dropOff: '',
            brutto: '',
            agent: '',
            remark: ''
        });

        const railForm = ref({
            origin: '',
            destination: '',
            coc: '',
            containerOwnership: 'no',
            hazard: 'no',
            security: 'no',
            profit: '',
            sum: '',
            agent: ''
        });

        const combDestinations = ref([]);
        const combForm = ref({
            seaPol: '',
            seaPod: '',
            dropOff: '',
            destination: '',
            coc: '',
            containerOwnership: 'no',
            hazard: 'no',
            transshipmentPort: '',
            security: 'Нет',
            seaProfit: '',
            railProfit: '',
            remark: ''
        });

        // Результаты расчетов
        const seaResults = ref([]);
        const railResults = ref([]);
        const combResults = ref([]);

        // Загружаемые списки
        const seaPods = ref([]);
        const seaDropOffLocations = ref([]);
        const railDestinations = ref([]);
        const combDropOffs = ref([]);
        const transshipmentPorts = ref([]);

        // Опции для селектов
        const calcTypes = ref([
            { title: '-- Выберите --', value: '' },
            { title: 'Расчёт морских перевозок', value: 'sea' },
            { title: 'Расчёт ж/д перевозок', value: 'rail' },
            { title: 'Комбинированный маршрут', value: 'combined' }
        ]);

        // Данные из PHP для селектов
        const seaPorts = ref(<?= json_encode(array_map(function($port) {
            return ['title' => $port, 'value' => $port];
        }, $seaPortsForVue), JSON_UNESCAPED_UNICODE) ?>);

        const zhdStarts = ref(<?= json_encode(array_map(function($station) {
            return ['title' => $station, 'value' => $station];
        }, $zhdStartsForVue), JSON_UNESCAPED_UNICODE) ?>);

        const combStarts = ref(<?= json_encode(array_map(function($port) {
            return ['title' => $port, 'value' => $port];
        }, $combStartsForVue), JSON_UNESCAPED_UNICODE) ?>);

        const seaCocTypes = ref([
            { title: '20DC', value: '20DC' },
            { title: '40HC', value: '40HC' }
        ]);

        const railCocTypes = ref([
            { title: '20DC (<24t)', value: '20DC (<24t)' },
            { title: '20DC (24t-28t)', value: '20DC (24t-28t)' },
            { title: '40HC (28t)', value: '40HC (28t)' }
        ]);

        const combCocTypes = ref([
            { title: '20DC (<24t)', value: '20DC (<24t)' },
            { title: '20DC (24t-28t)', value: '20DC (24t-28t)' },
            { title: '40HC (28t)', value: '40HC (28t)' }
        ]);

        const hazardOptions = ref([
            { title: 'Нет', value: 'no' },
            { title: 'Да', value: 'yes' }
        ]);

        const securityOptions = ref([
            { title: 'Нет', value: 'no' },
            { title: '20 фут', value: '20' },
            { title: '40 фут', value: '40' }
        ]);

        const combSecurityOptions = ref([
            { title: 'Нет', value: 'no' },
            { title: '20 фут', value: '20' },
            { title: '40 фут', value: '40' }
        ]);

        const ownershipOptions = ref([
            { title: 'Не выбрано', value: 'no' },
            { title: 'COC', value: 'coc' },
            { title: 'SOC', value: 'soc' }
        ]);

        // Заголовки для таблиц результатов
        const seaResultHeaders = ref([
            { title: 'Порт отправления', key: 'sea_pol', sortable: true },
            { title: 'Порт прибытия', key: 'sea_pod', sortable: true },
            { title: 'DROP OFF LOCATION', key: 'sea_drop_off_location', sortable: true },
            { title: 'Тип контейнера', key: 'sea_coc', sortable: true },
            { title: 'Собственность контейнера', key: 'sea_container_ownership', sortable: true },
            { title: 'CAF (%)', key: 'sea_caf_percent', sortable: true },
            { title: 'Стоимость обычного груза, USD', key: 'cost_total_normal', sortable: true },
            { title: 'Надбавка за опасный груз, USD', key: 'cost_total_danger', sortable: true },
            { title: 'Агент', key: 'sea_agent', sortable: true },
            { title: 'Примечание', key: 'sea_remark', sortable: true }
        ]);

        const railResultHeaders = ref([
            { title: 'Станция отправления', key: 'rail_origin', sortable: true },
            { title: 'Пункт назначения', key: 'rail_destination', sortable: true },
            { title: 'Тип контейнера', key: 'rail_coc', sortable: true },
            { title: 'Собственность контейнера', key: 'rail_container_ownership', sortable: true },
            { title: 'Охрана', key: 'rail_security', sortable: true },
            { title: 'Стоимость обычного груза, RUB', key: 'cost_total_normal', sortable: true },
            { title: 'Надбавка за опасный груз, RUB', key: 'cost_total_danger', sortable: true },
            { title: 'Агент', key: 'rail_agent', sortable: true },
            { title: 'Комментарий', key: 'rail_remark', sortable: true },
        ]);

        const combResultHeaders = ref([
            { title: 'Морской порт отправления', key: 'comb_sea_pol', sortable: true },
            { title: 'Морской порт прибытия', key: 'comb_sea_pod', sortable: true },
            { title: 'DROP OFF LOCATION', key: 'comb_drop_off', sortable: true },
            { title: 'Станции отправления', key: 'comb_rail_start', sortable: true },
            { title: 'Станции назначения', key: 'comb_rail_dest', sortable: true },
            { title: 'Тип контейнера', key: 'comb_coc', sortable: true },
            { title: 'Собственность контейнера', key: 'comb_container_ownership', sortable: true },
            { title: 'Охрана', key: 'comb_security', sortable: true },
            { title: 'Стоимость обычного груза, USD/RUB', key: 'cost_total_normal_text', sortable: true },
            { title: 'Надбавка за опасный груз, USD/RUB', key: 'cost_total_danger_text', sortable: true },
            { title: 'Агент', key: 'comb_agent', sortable: true },
            { title: 'Комментарий', key: 'comb_remark', sortable: true }
        ]);

        // Computed свойства
        const seaFormDisabled = computed(() => {
            return {
                pod: !seaForm.value.pol,
                coc: !seaForm.value.pod,
                dropOffLocation: !seaForm.value.coc,
                // Остальные поля зависят от выбора DROP OFF LOCATION
                hazard: !seaForm.value.dropOffLocation,
                security: !seaForm.value.dropOffLocation,
                containerOwnership: !seaForm.value.dropOffLocation,
                caf: !seaForm.value.dropOffLocation,
                profit: !seaForm.value.dropOffLocation
            };
        });

        const railFormDisabled = computed(() => {
            return {
                destination: !railForm.value.origin,
                coc: !railForm.value.destination,
                containerOwnership: !(railForm.value.origin && railForm.value.destination && railForm.value.coc),
                hazard: !(railForm.value.origin && railForm.value.destination && railForm.value.coc),
                security: !(railForm.value.origin && railForm.value.destination && railForm.value.coc),
                profit: !(railForm.value.origin && railForm.value.destination && railForm.value.coc)
            };
        });

        const combFormDisabled = computed(() => {
            return {
                dropOff: !combForm.value.seaPol,
                destination: !combForm.value.dropOff,
                coc: !combForm.value.dropOff,
                containerOwnership: !(combForm.value.seaPol && combForm.value.dropOff && combForm.value.coc),
                hazard: !(combForm.value.seaPol && combForm.value.dropOff),
                transshipmentPort: !combForm.value.dropOff,
                security: !(combForm.value.seaPol && combForm.value.dropOff),
                seaProfit: !(combForm.value.seaPol && combForm.value.dropOff && combForm.value.coc),
                railProfit: !(combForm.value.seaPol && combForm.value.dropOff && combForm.value.coc)
            };
        });

        const canCalculateSea = computed(() => {
            return seaForm.value.pol && seaForm.value.pod && seaForm.value.coc;
        });

        const canCalculateRail = computed(() => {
            return railForm.value.origin && railForm.value.destination && railForm.value.coc;
        });

        const canCalculateCombined = computed(() => {
            return combForm.value.seaPol && combForm.value.dropOff && combForm.value.coc;
        });

        // Методы
        const showLoading = (message) => {
            loadingMessage.value = message;
            loadingProgress.value = 0;
            loading.value = true;
        };

        const hideLoading = () => {
            loading.value = false;
            loadingProgress.value = 0;
        };

        const updateProgress = (percent, message = '') => {
            loadingProgress.value = Math.min(100, Math.max(0, percent));
            if (message) {
                loadingMessage.value = message;
            }
        };

        const showUploadMessage = (text, type = 'info') => {
            uploadMessage.value = text;
            uploadMessageType.value = type;
        };

        const clearUploadMessage = () => {
            uploadMessage.value = '';
        };

        // Морские перевозки
        const onSeaPolChange = () => {
            seaForm.value.pod = '';
            seaForm.value.coc = '';
            seaForm.value.dropOffLocation = '';
            seaForm.value.hazard = 'no';
            seaForm.value.security = 'no';
            seaForm.value.containerOwnership = 'no';
            seaForm.value.caf = '';
            seaForm.value.profit = '';
            seaForm.value.netto = '';
            seaForm.value.dropOff = '';
            seaForm.value.brutto = '';
            
            if (!seaForm.value.pol) {
                seaPods.value = [];
                return;
            }

            // Получаем уникальные POD для выбранного POL
            const pods = [...new Set(seaPerevozki
                .filter(r => r.POL === seaForm.value.pol)
                .map(r => r.POD))];
            seaPods.value = pods.map(pod => ({ title: pod, value: pod }));
        };

        const onSeaCocChange = () => {
            seaForm.value.dropOffLocation = '';
            seaForm.value.caf = '';
            seaForm.value.netto = '';
            seaForm.value.dropOff = '';
            seaForm.value.brutto = '';
            
            if (!seaForm.value.pod) {
                seaDropOffLocations.value = [];
                return;
            }

            // Получаем уникальные DROP_OFF_LOCATION для выбранных POL, POD и типа контейнера
            const locs = [...new Set(seaPerevozki
                .filter(r => 
                    r.POL === seaForm.value.pol && 
                    r.POD === seaForm.value.pod
                    // Фильтруем по наличию данных для выбранного типа контейнера
                    && (seaForm.value.coc === '20DC' 
                        ? (r.COC_20GP || r.SOC_20GP || r.DROP_OFF_20GP)
                        : (r.COC_40HC || r.SOC_40HC || r.DROP_OFF_40HC))
                )
                .map(r => r.DROP_OFF_LOCATION))];
            seaDropOffLocations.value = locs.map(loc => ({ title: loc, value: loc }));
        };

        const onSeaPodChange = () => {
            seaForm.value.caf = '';
            seaForm.value.netto = '';
            seaForm.value.dropOff = '';
            seaForm.value.brutto = '';
            
            if (!seaForm.value.pod) {
                seaDropOffLocations.value = [];
                return;
            }

            // Получаем уникальные DROP_OFF_LOCATION
            const locs = [...new Set(seaPerevozki
                .filter(r => r.POL === seaForm.value.pol && r.POD === seaForm.value.pod)
                .map(r => r.DROP_OFF_LOCATION))];
            seaDropOffLocations.value = locs.map(loc => ({ title: loc, value: loc }));
        };

        const onSeaDropOffChange = () => {
            seaForm.value.hazard = 'no';
            seaForm.value.security = 'no';
            seaForm.value.containerOwnership = 'no';
            seaForm.value.caf = '';
            seaForm.value.profit = '';
            seaForm.value.netto = '';
            seaForm.value.dropOff = '';
            seaForm.value.brutto = '';
            
            const pol = seaForm.value.pol;
            const pod = seaForm.value.pod;
            const loc = seaForm.value.dropOffLocation;
            const coc = seaForm.value.coc;
            
            if (!(pol && pod && loc && coc)) return;

            // Находим запись для заполнения полей
            const rec = seaPerevozki.find(r =>
                r.POL === pol &&
                r.POD === pod &&
                r.DROP_OFF_LOCATION === loc
            );
            
            if (rec) {
                seaForm.value.caf = rec.CAF_KONVERT || '';
                seaForm.value.remark = rec.REMARK || '';
                seaForm.value.agent = rec.AGENT || '';
            }
        };

        const updateSeaCosts = () => {
            const pol = seaForm.value.pol;
            const pod = seaForm.value.pod;
            const loc = seaForm.value.dropOffLocation;
            const coc = seaForm.value.coc;
            const ownership = seaForm.value.containerOwnership;
            
            seaForm.value.netto = '';
            seaForm.value.dropOff = '';
            
            if (!(pol && pod && loc && coc && ownership !== 'no')) {
                updateSeaBrutto();
                return;
            }

            // Находим запись
            const rec = seaPerevozki.find(r =>
                r.POL === pol &&
                r.POD === pod &&
                r.DROP_OFF_LOCATION === loc
            );
            
            if (!rec) {
                updateSeaBrutto();
                return;
            }

            // Получаем стоимость контейнера
            const ownershipKey = ownership.toUpperCase();
            let costCOC = 0;
            let dropOff = 0;
            
            if (coc === '20DC' && rec[`${ownershipKey}_20GP`]) {
                costCOC = parseFloat(rec[`${ownershipKey}_20GP`]) || 0;
                dropOff = parseFloat(rec.DROP_OFF_20GP) || 0;
            } else if (coc === '40HC' && rec[`${ownershipKey}_40HC`]) {
                costCOC = parseFloat(rec[`${ownershipKey}_40HC`]) || 0;
                dropOff = parseFloat(rec.DROP_OFF_40HC) || 0;
            }

            const caf = parseFloat(seaForm.value.caf) || 0;

            seaForm.value.dropOff = dropOff;
            const netto = (costCOC + dropOff);
            seaForm.value.netto = netto;

            updateSeaBrutto();
        };

        const updateSeaBrutto = () => {
            const netto = parseFloat(seaForm.value.netto) || 0;
            const profit = parseFloat(seaForm.value.profit) || 0;
            const caf = parseFloat(seaForm.value.caf) || 0;
            const dropOff = parseFloat(seaForm.value.dropOff) || 0;
            seaForm.value.brutto = ((dropOff + netto) * (caf / 100 + 1) + profit).toFixed(2);
        };

        const calculateSea = async () => {
            if (!canCalculateSea.value) return;
            
            showLoading('Расчет морских перевозок...');
            
            try {
                const payload = {
                    sea_pol: seaForm.value.pol,
                    sea_pod: seaForm.value.pod,
                    sea_drop_off_location: seaForm.value.dropOffLocation,
                    sea_coc: seaForm.value.coc,
                    sea_container_ownership: seaForm.value.containerOwnership,
                    sea_hazard: seaForm.value.hazard,
                    sea_caf: seaForm.value.caf,
                    sea_profit: seaForm.value.profit
                };

                const response = await fetch(cleanUrl + '?action=getSeaPerevozki', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                    body: new URLSearchParams(payload)
                });

                const data = await response.json();

                if (data.error) {
                    showUploadMessage(data.message || 'Ошибка расчета', 'error');
                    return;
                }

                seaResults.value = Array.isArray(data) ? data : [];
                showUploadMessage('Расчет завершен успешно!', 'success');
            } catch (error) {
                console.error('Ошибка расчета:', error);
                showUploadMessage('Ошибка при расчете: ' + error.message, 'error');
            } finally {
                hideLoading();
            }
        };

        // Ж/Д перевозки
        const onRailOriginChange = () => {
            railForm.value.destination = '';
            railForm.value.coc = '';
            railForm.value.containerOwnership = 'no';
            railForm.value.hazard = 'no';
            railForm.value.security = 'no';
            railForm.value.profit = '';
            railForm.value.sum = '';
            
            if (!railForm.value.origin) {
                railDestinations.value = [];
                return;
            }

            // Получаем уникальные станции назначения
            const dests = [...new Set(zhdPerevozki
                .filter(r => r.POL === railForm.value.origin)
                .map(r => r.POD))];
            railDestinations.value = dests.map(dest => ({ title: dest, value: dest }));
        };

        const onRailDestinationChange = () => {
            railForm.value.coc = '';
            railForm.value.containerOwnership = 'no';
            railForm.value.hazard = 'no';
            railForm.value.security = 'no';
            railForm.value.profit = '';
            railForm.value.sum = '';
            
            if (!railForm.value.destination) {
                railForm.value.agent = '';
                return;
            }

            // Находим запись для получения агента
            const origin = railForm.value.origin;
            const dest = railForm.value.destination;
            const rec = zhdPerevozki.find(r => r.POL === origin && r.POD === dest);
            
            if (rec) {
                railForm.value.agent = rec.AGENT || '';
            } else {
                railForm.value.agent = '';
            }
        };

        const calculateRail = async () => {
            if (!canCalculateRail.value) return;
            
            showLoading('Расчет ж/д перевозок...');
            
            try {
                const payload = {
                    rail_origin: railForm.value.origin,
                    rail_destination: railForm.value.destination,
                    rail_coc: railForm.value.coc,
                    rail_container_ownership: railForm.value.containerOwnership,
                    rail_hazard: railForm.value.hazard,
                    rail_security: railForm.value.security,
                    rail_profit: railForm.value.profit
                };

                const response = await fetch(cleanUrl + '?action=getRailPerevozki', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                    body: new URLSearchParams(payload)
                });

                const data = await response.json();

                if (data.error) {
                    showUploadMessage(data.message || 'Ошибка расчета', 'error');
                    return;
                }

                railResults.value = Array.isArray(data) ? data : [];
                showUploadMessage('Расчет завершен успешно!', 'success');
            } catch (error) {
                console.error('Ошибка расчета:', error);
                showUploadMessage('Ошибка при расчете: ' + error.message, 'error');
            } finally {
                hideLoading();
            }
        };

// Комбинированные перевозки
const onCombPolChange = () => {
    combForm.value.dropOff = '';
    combForm.value.destination = '';
    combForm.value.coc = '';
    combForm.value.containerOwnership = 'no';
    combForm.value.hazard = 'no';
    combForm.value.transshipmentPort = '';
    combForm.value.security = 'no';
    combForm.value.seaProfit = '';
    combForm.value.railProfit = '';
    combForm.value.remark = '';
    
    if (!combForm.value.seaPol) {
        combDropOffs.value = [];
        transshipmentPorts.value = [];
        combDestinations.value = [];
        return;
    }

    // Получаем уникальные DROP_OFF_LOCATION для выбранного POL
    const dropOffs = [...new Set(seaPerevozki
        .filter(item => item.POL === combForm.value.seaPol)
        .map(item => item.DROP_OFF_LOCATION))];
    combDropOffs.value = dropOffs.map(loc => ({ title: loc, value: loc }));
    
    transshipmentPorts.value = [];
    combDestinations.value = [];
    
    // Загружаем ВСЕ пункты назначения для выбранного POL
    loadAllDestinationsForPol();
};

// Новый метод для загрузки всех пунктов назначения для выбранного POL
const loadAllDestinationsForPol = () => {

    if (!combForm.value.seaPol) {
        combDestinations.value = [];
        return;
    }
    
    // Получаем уникальные POD (порты прибытия в морских перевозках)
    const seaPods = [...new Set(combPerevozki.map(item => item.POL))];

    // Получаем ВСЕ пункты назначения из комбинированных перевозок,
    // где POL соответствует любому из найденных POD
    const allDestinations = [...new Set(
        combPerevozki
            .map(item => item.PUNKT_NAZNACHENIYA)
            .filter(dest => dest && dest.trim() !== '')
    )];

    combDestinations.value = allDestinations.map(dest => ({ 
        title: dest, 
        value: dest 
    }));
};


const onCombDropOffChange = () => {
    combForm.value.destination = '';
    combForm.value.coc = '';
    combForm.value.containerOwnership = 'no';
    combForm.value.hazard = 'no';
    combForm.value.transshipmentPort = '';
    combForm.value.security = 'no';
    combForm.value.seaProfit = '';
    combForm.value.railProfit = '';
    combForm.value.remark = '';

    if (!combForm.value.dropOff) {
        transshipmentPorts.value = [];
        combDestinations.value = [];
        return;
    }

    const selectedPol = combForm.value.seaPol;
    const selectedDropOff = combForm.value.dropOff;

    // Получаем все POD для выбранных POL и DROP_OFF_LOCATION из морских перевозок
    const seaRecords = seaPerevozki.filter(item => 
        item.POL === selectedPol && 
        item.DROP_OFF_LOCATION === selectedDropOff
    );
    
    if (seaRecords.length === 0) {
        transshipmentPorts.value = [];
        combDestinations.value = [];
        return;
    }

    // Получаем уникальные POD (порты прибытия в морских перевозках)
    const seaPods = [...new Set(seaRecords.map(item => item.DROP_OFF_LOCATION))];

    // Получаем порты перевалки из комбинированных перевозок
    // Где POL в комбинированных перевозках соответствует POD из морских перевозок
    const transshipmentPoints = [...new Set(
        combPerevozki
            .filter(item => seaPods.includes(item.PUNKT_NAZNACHENIYA))
            .map(item => item.PUNKT_OTPRAVLENIYA)
            //.filter(point => point && point.trim() !== '')
    )];

    transshipmentPorts.value = transshipmentPoints.map(point => ({ 
        title: point, 
        value: point 
    }));

};

const onTransshipmentPortChange = () => {
    combForm.value.destination = '';
    combResults.value = []; // Очищаем результаты при изменении порта перевалки
    
    if (!combForm.value.transshipmentPort) {
        // Если порт перевалки не выбран, показываем все пункты назначения
        loadAllDestinationsForPol();
    } else {
        // Если порт перевалки выбран, получаем пункты назначения для этого порта
        const selectedPol = combForm.value.seaPol;
        const selectedDropOff = combForm.value.dropOff;
        const selectedTransshipment = combForm.value.transshipmentPort;
        
        // Получаем все POD для выбранных POL и DROP_OFF_LOCATION
        const seaRecords = seaPerevozki.filter(item => 
            item.POL === selectedPol && 
            item.DROP_OFF_LOCATION === selectedDropOff
        );
        
        if (seaRecords.length === 0) {
            combDestinations.value = [];
            return;
        }
        
        // Получаем уникальные POD из морских перевозок
        const seaPods = [...new Set(seaRecords.map(item => item.POD))];
        console.log(combPerevozki);
        // Получаем пункты назначения из комбинированных перевозок где:
        // 1. POL = POD из морских перевозок
        // 2. PUNKT_OTPRAVLENIYA = выбранный порт перевалки
        const destinations = [...new Set(
            combPerevozki
                .filter(item => 
                    //seaPods.includes(item.POL) && 
                    item.PUNKT_OTPRAVLENIYA === selectedTransshipment
                )
                .map(item => item.PUNKT_NAZNACHENIYA)
                .filter(dest => dest && dest.trim() !== '')
        )];
        
        combDestinations.value = destinations.map(dest => ({ 
            title: dest, 
            value: dest 
        }));
    }
};
const onCombDestinationChange = () => {
    if (!combForm.value.destination) {
        return;
    }
    
    // Здесь можно добавить дополнительную логику, если нужно
    // Например, фильтрацию станций назначения из ж/д перевозок
};
const calculateCombined = async () => {
    if (!canCalculateCombined.value) return;
    
    showLoading('Расчет комбинированного маршрута...');
    try {
        const payload = {
            comb_sea_pol: combForm.value.seaPol,
            comb_drop_off: combForm.value.dropOff,
            comb_rail_dest: combForm.value.destination || '',
            comb_coc: combForm.value.coc,
            comb_container_ownership: combForm.value.containerOwnership,
            comb_hazard: combForm.value.hazard,
            comb_transshipment_port: combForm.value.transshipmentPort || '',
            comb_security: combForm.value.security,
            sea_profit: combForm.value.seaProfit,
            rail_profit: combForm.value.railProfit
        };

        const response = await fetch(cleanUrl + '?action=getCombPerevozki', {
            method: 'POST',
            headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
            body: new URLSearchParams(payload)
        });

        const data = await response.json();
        
        if (data.error) {
            showUploadMessage(data.message || 'Ошибка расчета', 'error');
            return;
        }

        // Просто устанавливаем результаты без дополнительной фильтрации
        // так как фильтрация теперь происходит на сервере
        combResults.value = Array.isArray(data) ? data : [];
        
        if (combResults.value.length === 0) {
            showUploadMessage(
                'Нет результатов для выбранных параметров',
                'warning'
            );
        } else {
            showUploadMessage('Расчет завершен успешно! Найдено результатов: ' + combResults.value.length, 'success');
        }
    } catch (error) {
        console.error('Ошибка расчета:', error);
        showUploadMessage('Ошибка при расчете: ' + error.message, 'error');
    } finally {
        hideLoading();
    }
};
        // Обработка изменения типа расчета
        const onCalcTypeChange = () => {
            // Очищаем результаты при смене типа расчета
            seaResults.value = [];
            railResults.value = [];
            combResults.value = [];
            clearUploadMessage();
        };

        // Экспорт в Excel
const exportToExcel = async (type) => {
    let data = [];
    let headers = [];
    let filename = '';
    
    switch(type) {
        case 'sea':
            data = seaResults.value;
            headers = seaResultHeaders.value;
            filename = `морские_перевозки_${new Date().toISOString().slice(0, 19).replace(/:/g, '-')}.xlsx`;
            break;
        case 'rail':
            data = railResults.value;
            headers = railResultHeaders.value;
            filename = `жд_перевозки_${new Date().toISOString().slice(0, 19).replace(/:/g, '-')}.xlsx`;
            break;
        case 'combined':
            data = combResults.value;
            headers = combResultHeaders.value;
            filename = `комбинированные_перевозки_${new Date().toISOString().slice(0, 19).replace(/:/g, '-')}.xlsx`;
            break;
        default:
            return;
    }
    
    if (!data || data.length === 0) {
        showUploadMessage('Нет данных для экспорта', 'warning');
        return;
    }
    
    showLoading('Подготовка Excel файла...');
    
    try {
        // Создаем новую рабочую книгу
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Результаты');
        
        // Добавляем заголовки
        const headerRow = worksheet.addRow(headers.map(h => h.title));
        headerRow.font = { bold: true };
        headerRow.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFE0E0E0' }
        };
        
        // Добавляем данные
        data.forEach(item => {
            const rowData = headers.map(header => item[header.key] || '');
            worksheet.addRow(rowData);
        });
        
        // Настраиваем ширину столбцов
        worksheet.columns = headers.map((header, index) => {
            const headerLength = header.title.length;
            const maxDataLength = Math.max(
                ...data.map(row => {
                    const value = row[header.key];
                    return value ? value.toString().length : 0;
                })
            );
            return { 
                width: Math.max(headerLength, maxDataLength, 10) + 2,
                header: header.title
            };
        });
        
        // Добавляем границы ко всем ячейкам
        worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
            row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                cell.border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };
            });
        });
        
        // Замораживаем заголовок
        worksheet.views = [
            { state: 'frozen', xSplit: 0, ySplit: 1 }
        ];
        
        // Генерируем и скачиваем файл
        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { 
            type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
        });
        saveAs(blob, filename);
        
        showUploadMessage(`Файл "${filename}" успешно создан`, 'success');
    } catch (error) {
        console.error('Ошибка создания Excel:', error);
        showUploadMessage('Ошибка при создании Excel: ' + error.message, 'error');
    } finally {
        hideLoading();
    }
};
        // Загрузка файлов
        const fileInput = ref(null);
        const currentUploadType = ref('');

        const openFileDialog = (type) => {
            currentUploadType.value = type;
            fileInput.value.click();
        };

        const onFileSelected = async (event) => {
            const file = event.target.files[0];
            if (!file) return;

            const uploadMap = {
                'zhd': 'uploadZhd',
                'sea': 'uploadSea',
                'comb': 'uploadComb'
            };

            const action = uploadMap[currentUploadType.value];
            if (!action) return;

            // Показываем индикатор загрузки
            uploading.value = true;
            showLoading(`Загрузка файла ${file.name}...`);
            updateProgress(30, 'Отправка файла на сервер...');

            const formData = new FormData();
            formData.append('file', file);

            try {
                const response = await fetch(`${cleanUrl}?action=${encodeURIComponent(action)}`, {
                    method: 'POST',
                    body: formData,
                    credentials: 'same-origin'
                });

                updateProgress(70, 'Обработка файла на сервере...');
                
                const text = await response.text();

                let payload;
                try {
                    payload = text ? JSON.parse(text) : {};
                } catch {
                    payload = { error: text || response.statusText };
                }

                if (response.ok && (response.status === 200 || response.status === 207)) {
                    const added = payload.added ?? 0;
                    const errors = Array.isArray(payload.errors) ? payload.errors : [];
                    
                    if (response.status === 207 || errors.length) {
                        const firstErrors = errors.slice(0, 5).map(e => `строка ${e.row}: ${e.error}`).join('\n');
                        const more = errors.length > 5 ? `\n…и ещё ${errors.length - 5}` : '';
                        showUploadMessage(
                            `Загрузка завершена с частичными ошибками.\nДобавлено: ${added}.\nОшибки:\n${firstErrors}${more}`,
                            'warning'
                        );
                    } else {
                        showUploadMessage(payload.message || `Готово! Добавлено записей: ${added}.`, 'success');
                    }
                } else {
                    const errText = payload?.error || payload?.message || response.statusText || 'Неизвестная ошибка';
                    showUploadMessage(`Ошибка ${response.status}: ${errText}`, 'error');
                }
            } catch (error) {
                console.error('Ошибка загрузки:', error);
                showUploadMessage(`Сетевая ошибка: ${error.message || error}`, 'error');
            } finally {
                uploading.value = false;
                hideLoading();
                currentUploadType.value = '';
                event.target.value = ''; // Сброс input
            }
        };

        return {
            // Реактивные данные
            calcType,
            loading,
            loadingMessage,
            loadingProgress,
            uploading,
            uploadMessage,
            uploadMessageType,
            seaForm,
            railForm,
            combForm,
            seaResults,
            railResults,
            combResults,
            seaPorts,
            zhdStarts,
            combStarts,
            seaPods,
            seaDropOffLocations,
            seaCocTypes,
            railDestinations,
            railCocTypes,
            combDropOffs,
            combCocTypes,
            transshipmentPorts,
            seaFormDisabled,
            railFormDisabled,
            combFormDisabled,
            combDestinations,

            // Опции
            calcTypes,
            hazardOptions,
            securityOptions,
            combSecurityOptions,
            ownershipOptions,
            
            // Заголовки таблиц
            seaResultHeaders,
            railResultHeaders,
            combResultHeaders,
            
            // Computed свойства
            canCalculateSea,
            canCalculateRail,
            canCalculateCombined,
            
            // Методы
            showLoading,
            hideLoading,
            updateProgress,
            showUploadMessage,
            clearUploadMessage,
            onSeaPolChange,
            onSeaPodChange,
            onSeaDropOffChange,
            updateSeaCosts,
            updateSeaBrutto,
            calculateSea,
            onRailOriginChange,
            onRailDestinationChange,
            calculateRail,
            onCombPolChange,
            onCombDropOffChange,
            calculateCombined,
            onCalcTypeChange,
            exportToExcel,
            openFileDialog,
            onFileSelected,
            fileInput,
            onSeaCocChange,
            onTransshipmentPortChange,
        };
    }
}).use(Vuetify.createVuetify()).mount('#app');
</script>
</body>
</html>