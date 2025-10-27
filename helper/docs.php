<?php

if (! function_exists('docs')) {
    function docs(...$arguments)
    {
        if (! isset($GLOBALS['DOC_GENERATOR'])) {
            return;
        }

        if (count($arguments) > 2) {
            $in = $arguments[0] ?? [];
            $process = $arguments[1] ?? [];
            $out = $arguments[2] ?? [];
            $validate = $arguments[3] ?? [];
        } elseif (count($arguments) > 1) {
            $in = $arguments[0] ?? [];
            $process = $arguments[1] ?? [];
            $out = [];
            $validate = [];
        } else {
            $in = [];
            $process = $arguments[0] ?? [];
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
