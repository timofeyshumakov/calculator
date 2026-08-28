<?php

declare(strict_types=1);

namespace Calculator\Http;

final class JsonResponse
{
    public static function send(mixed $payload, int $status = 200): void
    {
        http_response_code($status);
        header('Content-Type: application/json; charset=utf-8');
        echo json_encode($payload, JSON_UNESCAPED_UNICODE);
    }

    public static function error(string $message, int $status = 400, bool $asErrorFlag = true): void
    {
        self::send([
            'error' => $asErrorFlag,
            'message' => $message,
        ], $status);
    }
}
