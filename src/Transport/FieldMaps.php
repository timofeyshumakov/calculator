<?php

declare(strict_types=1);

namespace Calculator\Transport;

/**
 * Маппинги полей инфоблоков Bitrix → внутренние ключи калькулятора.
 */
final class FieldMaps
{
    public const IBLOCK_RAIL = 30;
    public const IBLOCK_SEA = 28;
    public const IBLOCK_COMBINED = 32;

    public const RAIL = [
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
        'PROPERTY_228' => 'PORT_PEREVALKI',
    ];

    public const SEA = [
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

    public const COMBINED = [
        'NAME' => 'POL',
        'PROPERTY_182' => 'PUNKT_OTPRAVLENIYA',
        'PROPERTY_184' => 'STANTSIYA_OTPRAVLENIYA',
        'PROPERTY_186' => 'PUNKT_NAZNACHENIYA',
        'PROPERTY_188' => 'STANTSIYA_NAZNACHENIYA',
        'PROPERTY_190' => 'REMARK',
    ];
}
