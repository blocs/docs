<?php

if (! function_exists('docs')) {
    function docs(...$arguments)
    {
        if (! isset($GLOBALS['DOC_GENERATOR'])) {
            return;
        }

        // 引数が1つなら処理だけ、2つなら「入力, 処理」、3つ以上なら「入力, 処理, 出力, バリデーション」
        [$inputDocs, $processDocs, $outputDocs, $validationDocs] = match (true) {
            count($arguments) > 2 => $arguments + [3 => null],
            count($arguments) > 1 => [$arguments[0], $arguments[1], null, null],
            default => [null, $arguments[0] ?? null, null, null],
        };

        // 配列でない指定は配列にそろえる（入力・出力・バリデーションは 名前 => 詳細、処理は処理のリスト）
        $normalizeDocs = static function ($value, bool $associative) {
            if (is_array($value)) {
                return $value;
            }
            if ($value === null) {
                return [];
            }

            if (! $associative) {
                return [$value];
            }

            // 配列のキーにできない値（オブジェクトなど）は、記録しない
            return is_string($value) || is_int($value) || is_float($value) || is_bool($value) ? [(string) $value => []] : [];
        };

        $inputDocs = $normalizeDocs($inputDocs, true);
        $processDocs = $normalizeDocs($processDocs, false);
        $outputDocs = $normalizeDocs($outputDocs, true);
        $validationDocs = $normalizeDocs($validationDocs, true);

        $backtrace = debug_backtrace(DEBUG_BACKTRACE_IGNORE_ARGS, 2);

        $GLOBALS['DOC_GENERATOR'][] = [
            'path' => $backtrace[0]['file'] ?? '',
            'function' => $backtrace[1]['function'] ?? '',
            'line' => $backtrace[0]['line'] ?? 0,
            'in' => $inputDocs,
            'process' => $processDocs,
            'out' => $outputDocs,
            'validate' => $validationDocs,
        ];
    }
}
