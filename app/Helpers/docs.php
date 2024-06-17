<?php

if (!function_exists('docs')) {
    function docs(...$argvs)
    {
        if (!isset($GLOBALS['DOC_GENERATOR'])) {
            return;
        }

        if (count($argvs) > 2) {
            $in = $argvs[0] ?? [];
            $process = $argvs[1] ?? [];
            $out = $argvs[2] ?? [];
            $validate = $argvs[3] ?? [];
        } elseif (count($argvs) > 1) {
            $in = $argvs[0] ?? [];
            $process = $argvs[1] ?? [];
            $out = [];
            $validate = [];
        } else {
            $in = [];
            $process = $argvs[0] ?? [];
            $out = [];
            $validate = [];
        }

        is_array($in) || $in = [$in => []];
        is_array($process) || $process = [$process];
        is_array($out) || $out = [$out => []];
        is_array($validate) || $validate = [$validate => []];

        $backtrace = debug_backtrace();
        $GLOBALS['DOC_GENERATOR'][] = [
            'path' => $backtrace[0]['file'],
            'function' => $backtrace[1]['function'],
            'line' => $backtrace[0]['line'],
            'in' => $in,
            'process' => $process,
            'out' => $out,
            'validate' => $validate,
        ];
    }
}
