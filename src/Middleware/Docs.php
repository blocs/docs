<?php

namespace Blocs\Middleware;

use Blocs\Excel;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\Route;
use Symfony\Component\HttpFoundation\Response;

class Docs
{
    private array $keywords = [];

    private array $neglectPatterns = [];

    private array $commentMap = [];

    public function handle(Request $request, \Closure $next): Response
    {
        // ドキュメント生成で利用するグローバル情報を初期化
        $GLOBALS['DOC_GENERATOR'] = [];

        $response = $next($request);

        if (! file_exists(base_path('docs/format.xlsx'))) {
            return $response;
        }

        // 現在のコントローラーとメソッドを判定
        $currentRouteAction = ltrim(str_replace('\\', '/', Route::currentRouteAction()), '/');
        $currentRouteAction = str_replace('App/Http/Controllers/', '', $currentRouteAction);
        empty($currentRouteAction) && $currentRouteAction = 'class@method';
        [$routeClass, $routeMethod] = explode('@', $currentRouteAction, 2);

        // ドキュメント用のエクセルファイルを準備
        $excelPath = base_path("docs/{$currentRouteAction}.xlsx");
        is_dir(dirname($excelPath)) || mkdir(dirname($excelPath), 0777, true);
        copy(base_path('docs/format.xlsx'), $excelPath);
        $excel = new Excel($excelPath);

        // 設定ファイルを読み込み反映
        $this->loadConfig($routeClass, $routeMethod, $excel);

        $startLine = 5;
        $headlineNo = 1;
        $indentNo = 1;
        $steps = $GLOBALS['DOC_GENERATOR'];

        if (count($steps)) {
            $endNo = count($steps) - 1;

            if (! $steps[$endNo]['in'] && $response->getStatusCode() === 200 && is_object($response->original) && method_exists($response->original, 'getPath')) {
                // 画面描画の入力情報を補完
                $viewPath = str_replace(resource_path('views/'), '', $response->original->getPath());
                $viewPath && $steps[$endNo]['in'] = ['テンプレート' => '!'.$viewPath];
            }

            if (! $steps[$endNo]['out']) {
                // 画面描画の出力情報を補完
                if ($response->getStatusCode() === 200) {
                    $contents = str_replace(["\r\n", "\r", "\n"], '', $response->getContent());
                    if (preg_match('/<title>(.*?)<\/title>/i', $contents, $match)) {
                        $steps[$endNo]['out'] = ['HTML' => '!'.trim($match[1])];
                    }
                }
            }
        }

        foreach ($steps as $stepNo => $step) {
            // 非表示対象のステップを判定
            $stepProcess = implode('', $step['process']);
            $stepProcess = $this->normalizeProcessValue($stepProcess);
            if ($this->shouldSkipStep($stepProcess)) {
                continue;
            }

            $maxLine = $startLine;

            // 入力情報を記述
            $line = $this->fillInputRows($startLine, $step, $excel);
            $line > $maxLine && $maxLine = $line;

            // 処理手順を記述
            $line = $this->fillProcessRows($startLine, $step, $excel, $headlineNo, $indentNo);
            $line > $maxLine && $maxLine = $line;

            // 出力情報を記述
            $line = $this->fillOutputRows($startLine, $step, $excel);
            $line > $maxLine && $maxLine = $line;

            // 開始行更新
            $startLine = $maxLine;
        }

        $excel->name(1, $routeMethod)->save($excelPath);

        return $response;
    }

    private function fillInputRows($line, $step, $excel)
    {
        foreach ($step['in'] as $key => $items) {
            $excel->set(1, 'A', $line, $key);
            $excel->set(1, 'J', $line, '→');
            $line++;

            is_array($items) || $items = array_filter([$items], 'strlen');
            foreach ($items as $item) {
                $excel->set(1, 'B', $line, $this->normalizeInOutValue($item));
                $line++;
            }
        }

        return ++$line;
    }

    private function fillProcessRows($line, $step, $excel, &$headlineNo, &$indentNo)
    {
        $pathColumn = 'M';

        foreach ($step['process'] as $process) {
            $comments = explode("\n", $process);
            $process = array_shift($comments);

            // 行頭が#のときは見出し扱い
            $headline = ! strncmp($process, '#', 1);
            $headline && $process = trim(substr($process, 1));

            $column = $headline ? 'K' : 'L';
            $pathColumn = $headline ? 'L' : 'M';
            $process = $this->normalizeProcessValue($process);
            if ($headline) {
                // 見出し行を記述
                $excel->set(1, $column, $line, $headlineNo.'. '.$process);
                $headlineNo++;
                $indentNo = 1;
            } else {
                // 見出し配下の処理を記述
                $excel->set(1, $column, $line, $indentNo.') '.$process);
                $indentNo++;
            }
            $line++;

            // 追加コメントを補完
            $column = $headline ? 'L' : 'M';
            ($addComment = $this->findSupplementaryComment($process)) && $comments = array_merge($comments, explode("\n", $addComment));

            // バリデーション情報を整形
            count($step['validate']) && $comments[] .= '<入力値>: <条件>: <メッセージ>';
            foreach ($step['validate'] as $validate) {
                $validateComment = '・'.$validate['name'];
                empty($validate['validate']) || $validateComment .= ': '.$validate['validate'];
                empty($validate['message']) || $validateComment .= ': '.$validate['message'];
                $comments[] .= $validateComment;
            }

            foreach ($comments as $comment) {
                $excel->set(1, $column, $line, $this->normalizeProcessValue($comment));
                $line++;
            }
        }

        // 処理の箇所を記述
        $path = str_replace(base_path().'/', '', $step['path']);
        $excel->set(1, $pathColumn, $line, $path.'@'.$step['function'].':'.$step['line']);
        $line++;

        return ++$line;
    }

    private function fillOutputRows($line, $step, $excel)
    {
        foreach ($step['out'] as $key => $items) {
            $excel->set(1, 'AO', $line, '→');
            $excel->set(1, 'AP', $line, $key);
            $line++;

            is_array($items) || $items = array_filter([$items], 'strlen');
            foreach ($items as $item) {
                $excel->set(1, 'AQ', $line, $this->normalizeInOutValue($item));
                $line++;
            }
        }

        return ++$line;
    }

    private function loadConfig($routeClass, $routeMethod, $excel)
    {
        $config = [];
        $keywords = [];
        $neglectPatterns = [];
        $commentMap = [];

        $excel->set(1, 'AU', '1', date('Y/m/d'));
        $excel->set(1, 'E', '2', $routeClass.'@'.$routeMethod);

        if (file_exists(base_path('docs/common.php'))) {
            include base_path('docs/common.php');

            $keywords = $config['keyword'] ?? [];
            $neglectPatterns = $config['neglect'] ?? [];
            $commentMap = $config['comment'] ?? [];
        }

        if (file_exists(base_path('docs/'.$routeClass.'.php'))) {
            include base_path('docs/'.$routeClass.'.php');

            // classとmethodの概要を記述
            isset($config['description']) && $excel->set(1, 'Z', '1', $config['description']);
            isset($config[$routeMethod]['description']) && $excel->set(1, 'Q', '2', $config[$routeMethod]['description']);

            // キーワードを取得
            isset($config['keyword']) && $keywords = array_merge($keywords, $config['keyword']);
            isset($config[$routeMethod]['keyword']) && $keywords = array_merge($keywords, $config[$routeMethod]['keyword']);

            // 非表示行を取得
            isset($config['neglect']) && $neglectPatterns = array_merge($neglectPatterns, $config['neglect']);
            isset($config[$routeMethod]['neglect']) && $neglectPatterns = array_merge($neglectPatterns, $config[$routeMethod]['neglect']);

            // 追加コメントを取得
            isset($config['comment']) && $commentMap = $this->mergeConfig($commentMap, $config['comment']);
            isset($config[$routeMethod]['comment']) && $commentMap = $this->mergeConfig($commentMap, $config[$routeMethod]['comment']);
        }

        $this->keywords = $keywords;
        $this->neglectPatterns = $neglectPatterns;
        $this->commentMap = $commentMap;
    }

    private function mergeConfig(array $before, array $after): array
    {
        foreach ($after as $key => $value) {
            $before[$key] = $value;
        }

        return $before;
    }

    private function normalizeInOutValue($item)
    {
        if (! strncmp($item, '!', 1)) {
            return substr($item, 1);
        }

        // キーワードを長い順にソート
        $keywords = array_keys($this->keywords);
        array_multisort(array_map('strlen', $keywords), SORT_DESC, $keywords);

        foreach ($keywords as $key) {
            if (strpos($item, $key) !== false) {
                // キーワード置き換え
                $item = str_replace($key, $key.': '.$this->keywords[$key], $item);

                return $item;
            }
        }

        return $item;
    }

    private function normalizeProcessValue($item)
    {
        if (! strncmp($item, '!', 1)) {
            return substr($item, 1);
        }

        foreach ($this->keywords as $key => $value) {
            // キーワード置き換え
            $item = str_replace('<'.$key.'>', '<'.$value.'>', $item);
        }

        return $item;
    }

    private function shouldSkipStep($item)
    {
        $item = preg_replace("/\s/", '', $item);
        foreach ($this->neglectPatterns as $neglect) {
            $neglect = preg_replace("/\s/", '', $neglect);

            if (strpos($item, $neglect) !== false) {
                return true;
            }
        }

        return false;
    }

    private function findSupplementaryComment($item)
    {
        $item = preg_replace("/\s/", '', $item);
        $commentKeys = array_keys($this->commentMap);
        foreach ($commentKeys as $commentKey) {
            $commentKey = preg_replace("/\s/", '', $commentKey);

            if (strpos($item, $commentKey) !== false) {
                return $this->commentMap[$commentKey];
            }
        }

        return false;
    }
}
