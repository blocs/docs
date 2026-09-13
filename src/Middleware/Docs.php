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
        $currentRouteAction = self::normalizeRouteAction(Route::currentRouteAction());
        [$routeClass, $routeMethod] = explode('@', $currentRouteAction, 2);

        // ドキュメント用のエクセルファイルを準備
        $excelPath = base_path("docs/{$currentRouteAction}.xlsx");
        is_dir(dirname($excelPath)) || mkdir(dirname($excelPath), 0755, true);
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

            if (! $steps[$endNo]['in'] && $response->getStatusCode() === 200) {
                $original = $response instanceof \Illuminate\Http\Response
                    ? $response->getOriginalContent()
                    : null;
                if (is_object($original) && method_exists($original, 'getPath')) {
                    // 画面描画の入力情報を補完
                    $viewPath = str_replace(resource_path('views/'), '', $original->getPath());
                    $viewPath && $steps[$endNo]['in'] = ['テンプレート' => '!'.$viewPath];
                }
            }

            if (! $steps[$endNo]['out']) {
                // 画面描画の出力情報を補完
                if ($response->getStatusCode() === 200) {
                    $contents = $response->getContent();
                    if (is_string($contents) && $contents !== '') {
                        $contents = str_replace(["\r\n", "\r", "\n"], '', substr($contents, 0, 200000));
                        if (preg_match('/<title>(.*?)<\/title>/i', $contents, $match)) {
                            $steps[$endNo]['out'] = ['HTML' => '!'.trim($match[1])];
                        }
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
            $line = $this->fillIoRows($startLine, $step['in'], $excel, ['A' => null, 'J' => '→'], 'B');
            $line > $maxLine && $maxLine = $line;

            // 処理手順を記述
            $line = $this->fillProcessRows($startLine, $step, $excel, $headlineNo, $indentNo);
            $line > $maxLine && $maxLine = $line;

            // 出力情報を記述
            $line = $this->fillIoRows($startLine, $step['out'], $excel, ['AO' => '→', 'AP' => null], 'AQ');
            $line > $maxLine && $maxLine = $line;

            // 開始行更新
            $startLine = $maxLine;
        }

        $excel->name(1, $routeMethod)->save($excelPath);

        return $response;
    }

    /**
     * @param  array<string, string|null>  $header  列 => 固定値（null はキー名）
     */
    private function fillIoRows($line, array $rows, $excel, array $header, string $valueColumn)
    {
        foreach ($rows as $key => $items) {
            foreach ($header as $column => $value) {
                $excel->set(1, $column, $line, $value ?? $key);
            }
            $line++;

            is_array($items) || $items = array_filter([$items], 'strlen');
            foreach ($items as $item) {
                $excel->set(1, $valueColumn, $line, $this->normalizeInOutValue($item));
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
            $label = $headline
                ? $headlineNo.'. '.$process
                : $indentNo.') '.$process;
            if ($headline) {
                $headlineNo++;
                $indentNo = 1;
            } else {
                $indentNo++;
            }
            $excel->set(1, $column, $line, $label);
            $line++;

            // 追加コメントを補完
            $column = $headline ? 'L' : 'M';
            ($addComment = $this->findSupplementaryComment($process)) && $comments = array_merge($comments, explode("\n", $addComment));

            // バリデーション情報を整形
            count($step['validate']) && $comments[] = '<入力値>: <条件>: <メッセージ>';
            foreach ($step['validate'] as $validate) {
                $validateComment = '・'.$validate['name'];
                empty($validate['validate']) || $validateComment .= ': '.$validate['validate'];
                empty($validate['message']) || $validateComment .= ': '.$validate['message'];
                $comments[] = $validateComment;
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
            isset($config['comment']) && $commentMap = array_replace($commentMap, $config['comment']);
            isset($config[$routeMethod]['comment']) && $commentMap = array_replace($commentMap, $config[$routeMethod]['comment']);
        }

        $this->keywords = $keywords;
        $this->neglectPatterns = $neglectPatterns;
        $this->commentMap = $commentMap;
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
        $item = $this->stripWhitespace($item);
        foreach ($this->neglectPatterns as $neglect) {
            if (strpos($item, $this->stripWhitespace($neglect)) !== false) {
                return true;
            }
        }

        return false;
    }

    private function findSupplementaryComment($item)
    {
        $item = $this->stripWhitespace($item);
        foreach ($this->commentMap as $commentKey => $comment) {
            $normalizedKey = $this->stripWhitespace($commentKey);
            if ($normalizedKey !== '' && strpos($item, $normalizedKey) !== false) {
                return $comment;
            }
        }

        return false;
    }

    private function stripWhitespace($item): string
    {
        return preg_replace("/\s/", '', (string) $item);
    }

    /**
     * @internal テストから正規化結果を検証する
     */
    public static function normalizeRouteAction(mixed $currentRouteAction): string
    {
        if (! is_string($currentRouteAction) || $currentRouteAction === '') {
            return 'class@method';
        }

        $currentRouteAction = ltrim(str_replace('\\', '/', $currentRouteAction), '/');
        $currentRouteAction = str_replace('App/Http/Controllers/', '', $currentRouteAction);
        if ($currentRouteAction === '') {
            return 'class@method';
        }

        if (! str_contains($currentRouteAction, '@')) {
            $currentRouteAction .= '@__invoke';
        }

        return $currentRouteAction;
    }
}
