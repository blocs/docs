<?php

if (! function_exists('docs')) {
    function docs(...$arguments)
    {
        if (! isset($GLOBALS['DOC_GENERATOR'])) {
            return;
        }

        $argumentCount = count($arguments);

        $inputDocs = [];
        $processDocs = [];
        $outputDocs = [];
        $validationDocs = [];

        if ($argumentCount > 2) {
            $inputDocs = $arguments[0] ?? [];
            $processDocs = $arguments[1] ?? [];
            $outputDocs = $arguments[2] ?? [];
            $validationDocs = $arguments[3] ?? [];
        } elseif ($argumentCount > 1) {
            $inputDocs = $arguments[0] ?? [];
            $processDocs = $arguments[1] ?? [];
        } else {
            $processDocs = $arguments[0] ?? [];
        }

        $normalizeDocs = static function ($value, bool $associative) {
            if (is_array($value)) {
                return $value;
            }

            return $associative ? [$value => []] : [$value];
        };

        $inputDocs = $normalizeDocs($inputDocs, true);
        $processDocs = $normalizeDocs($processDocs, false);
        $outputDocs = $normalizeDocs($outputDocs, true);
        $validationDocs = $normalizeDocs($validationDocs, true);

        $backtrace = debug_backtrace(DEBUG_BACKTRACE_IGNORE_ARGS, 2);

        $GLOBALS['DOC_GENERATOR'][] = [
            'path' => $backtrace[0]['file'],
            'function' => $backtrace[1]['function'],
            'line' => $backtrace[0]['line'],
            'in' => $inputDocs,
            'process' => $processDocs,
            'out' => $outputDocs,
            'validate' => $validationDocs,
        ];
    }
}
