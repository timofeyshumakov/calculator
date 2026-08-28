💎 Калькулятор стоимости перевозок (море / ЖД / комбинированные) для Bitrix24.

## О проекте

Локальное приложение Bitrix24: расчёт ставок морских, железнодорожных и комбинированных перевозок по справочникам списков CRM.

## Стек

- PHP 8.1+
- Bitrix24 REST (CRest)
- Vue 3 + Vuetify 3 (UI в `Forms.php`)
- PHPUnit 10 (чистая логика расчёта)

## Структура (middle)

```
src/
  Calculator/     # SeaCalculator, RailCalculator, CombinedCalculator
  Support/        # CostMath, BitrixValueNormalizer
  Transport/      # FieldMaps, TransportCache
  Http/           # JsonResponse
tests/            # unit-тесты формул и маппинга
Controller.php    # HTTP-экшены + оркестрация
Forms.php         # UI калькулятора
demo/data/        # мок-справочники для локальной демонстрации
```

## Установка

```bash
composer install
```

Точка входа Bitrix: `Controller.php?action=index`  
Нужны `crestV136/` и `app_config.php` на сервере (в git не кладём — секреты).

## Тесты

```bash
composer test
# или
vendor/bin/phpunit
```

## Формулы (кратко)

- **Море:** `netto = ceil(ставка + drop-off)`, `total = ceil(netto * (1 + CAF%/100) + прибыль)`
- **ЖД:** `total = ceil(база + охрана + прибыль)`
- Округление вверх (`ceil`) для положительных ставок

## API (основные action)

| action | назначение |
|--------|------------|
| `index` | форма |
| `getTransportData` | справочники (sea/rail/comb) |
| `getSeaPerevozki` | расчёт моря |
| `getRailPerevozki` | расчёт ЖД |
| `getCombPerevozki` | комбинированный расчёт |
| `install` | установка приложения B24 |

## Лицензия

Proprietary.
