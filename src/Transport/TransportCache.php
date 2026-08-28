<?php

declare(strict_types=1);

namespace Calculator\Transport;

/**
 * Файловый кэш справочников перевозок (stale-while-revalidate).
 */
final class TransportCache
{
    public function __construct(
        private readonly string $directory,
        private readonly int $ttlSeconds = 21600,
    ) {
    }

    public function ensureDirectory(): string
    {
        if (!is_dir($this->directory)) {
            mkdir($this->directory, 0775, true);
        }

        return $this->directory;
    }

    public function path(string $cacheKey): string
    {
        return $this->ensureDirectory() . DIRECTORY_SEPARATOR . $cacheKey . '.json';
    }

    public function read(string $cacheKey, bool $allowStale = false): ?array
    {
        $cacheFile = $this->path($cacheKey);
        if (!file_exists($cacheFile)) {
            return null;
        }

        $isExpired = (time() - filemtime($cacheFile)) >= $this->ttlSeconds;
        if ($isExpired && !$allowStale) {
            return null;
        }

        $cached = json_decode((string) file_get_contents($cacheFile), true);
        return is_array($cached) ? $cached : null;
    }

    public function isFresh(string $cacheKey): bool
    {
        $cacheFile = $this->path($cacheKey);
        if (!file_exists($cacheFile)) {
            return false;
        }

        return (time() - filemtime($cacheFile)) < $this->ttlSeconds;
    }

    public function write(string $cacheKey, array $data): void
    {
        file_put_contents(
            $this->path($cacheKey),
            json_encode($data, JSON_UNESCAPED_UNICODE),
            LOCK_EX
        );
    }

    public function clear(?int $iblockId = null): void
    {
        $dir = $this->ensureDirectory();

        if ($iblockId !== null) {
            $cacheFile = $dir . DIRECTORY_SEPARATOR . 'transport_' . $iblockId . '.json';
            if (file_exists($cacheFile)) {
                unlink($cacheFile);
            }
            return;
        }

        foreach (glob($dir . DIRECTORY_SEPARATOR . 'transport_*.json') ?: [] as $cacheFile) {
            unlink($cacheFile);
        }
    }
}
