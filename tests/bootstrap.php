<?php

$packageRoot = dirname(__DIR__);
// 1) 環境変数 BLOCS_TEST_AUTOLOAD で任意のアプリの autoload.php を指定した場合
// 2) パッケージ単体で composer install した場合
// 3) アプリの vendor/blocs/docs として入っている場合
$autoloadCandidates = array_filter([
    getenv('BLOCS_TEST_AUTOLOAD') ?: null,
    $packageRoot.'/vendor/autoload.php',
    dirname($packageRoot, 2).'/autoload.php',
]);

foreach ($autoloadCandidates as $autoload) {
    if (is_file($autoload)) {
        require_once $autoload;
        break;
    }
}

spl_autoload_register(static function (string $class) use ($packageRoot): bool {
    $prefix = 'Blocs\\';
    if (! str_starts_with($class, $prefix)) {
        return false;
    }

    $file = $packageRoot.'/src/'.str_replace('\\', '/', substr($class, strlen($prefix))).'.php';
    if (is_file($file)) {
        require $file;

        return true;
    }

    return false;
}, true, true);
