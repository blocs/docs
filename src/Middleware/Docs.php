<?php

namespace Blocs\Middleware;

use Blocs\Excel;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\Route;
use Symfony\Component\HttpFoundation\Response;

class Docs
{
    private array $keywords = [];

    /** @var array<int, int|string> キーワードを長い順に並べたもの */
    private array $keywordsByLength = [];

    private array $neglectPatterns = [];

    private array $commentMap = [];

    public function handle(Request $request, \Closure $next): Response
    {
        // 同じリクエストにこのミドルウェアが二重に登録されている場合（グローバルとルートの両方など）は、
        // 最初に通った1回だけが生成する（もう一方が、記録のない空の設計書で上書きしないようにする）
        // （印はリクエストのオブジェクトごと。duplicate() で作ったリクエストには引き継がせず、終わったら消す）
        if (! $this->shouldGenerate() || $request->attributes->get(self::class) === spl_object_id($request)) {
            return $next($request);
        }
        $request->attributes->set(self::class, spl_object_id($request));

        // ドキュメント生成で利用するグローバル情報を初期化
        // （処理の中から別のリクエストを実行した場合に備えて、呼び出し元の記録は退避しておく）
        $outerSteps = $GLOBALS['DOC_GENERATOR'] ?? null;
        $GLOBALS['DOC_GENERATOR'] = [];

        try {
            $response = $next($request);

            try {
                $this->generateDocument($request, $response, $GLOBALS['DOC_GENERATOR'] ?? []);
            } catch (\Throwable $exception) {
                // ドキュメントの生成に失敗しても、本来のレスポンスは返す
                function_exists('report') && report($exception);
            }

            return $response;
        } finally {
            $request->attributes->remove(self::class);

            // 常駐ワーカーで、このミドルウェアを通らないリクエストの docs() が記録を溜め続けないようにする
            if ($outerSteps === null) {
                unset($GLOBALS['DOC_GENERATOR']);
            } else {
                $GLOBALS['DOC_GENERATOR'] = $outerSteps;
            }
        }
    }

    /**
     * ドキュメントを生成するか。雛形（docs/format.xlsx）があり、本番環境でない場合だけ生成する
     * （本番環境でリクエストのたびにファイルを書き出さない）
     */
    private function shouldGenerate(): bool
    {
        if (! file_exists(base_path('docs/format.xlsx'))) {
            return false;
        }

        return ! (function_exists('app') && method_exists(app(), 'isProduction') && app()->isProduction());
    }

    /**
     * @param  array<int, array<string, mixed>>  $steps  docs() で記録した処理
     */
    private function generateDocument(Request $request, Response $response, array $steps): void
    {
        // このリクエストのコントローラーとメソッドを判定
        // （処理の中から別のリクエストを実行すると、ルーターの「現在のルート」は内側のままになるため、
        //   ミドルウェアに渡されたリクエスト自身のルートから求める）
        $route = $request->route();
        $currentRouteAction = self::normalizeRouteAction(
            is_object($route) && method_exists($route, 'getAction') ? $route->getAction('controller') : Route::currentRouteAction()
        );
        [$routeClass, $routeMethod] = explode('@', $currentRouteAction, 2);

        // ドキュメント用のエクセルファイルを準備（雛形を読み込み、保存先へ書き出す）
        $excelPath = base_path("docs/{$currentRouteAction}.xlsx");
        // 同時アクセスで他のリクエストが先に作った場合も、失敗にしない
        is_dir(dirname($excelPath)) || @mkdir(dirname($excelPath), 0755, true) || is_dir(dirname($excelPath));
        $excel = new Excel(base_path('docs/format.xlsx'));

        // 設定ファイルを読み込み反映
        $this->loadConfig($routeClass, $routeMethod, $excel);

        $startLine = 5;
        $headlineNo = 1;
        $indentNo = 1;

        $steps = $this->completeLastStep($steps, $response);

        foreach ($steps as $step) {
            // 非表示対象のステップを判定
            $stepProcess = implode('', array_map([$this, 'stringify'], $step['process']));
            $stepProcess = $this->normalizeProcessValue($stepProcess);
            if ($this->shouldSkipStep($stepProcess)) {
                continue;
            }

            // 入力情報・処理手順・出力情報を横に並べて記述し、一番下まで使った行の次から次のステップを始める
            $startLine = max(
                $this->fillIoRows($startLine, $step['in'], $excel, ['A' => null, 'J' => '→'], 'B'),
                $this->fillProcessRows($startLine, $step, $excel, $headlineNo, $indentNo),
                $this->fillIoRows($startLine, $step['out'], $excel, ['AO' => '→', 'AP' => null], 'AQ')
            );
        }

        // save() は一時ファイルに書いてから置き換えるので、同じ画面への同時アクセスでも壊れたファイルは残らない
        if (! $excel->name(1, $routeMethod)->save($excelPath)) {
            // handle() で捕捉され、ログに記録される
            throw new \RuntimeException('Can not save document: '.$excelPath);
        }
    }

    /**
     * 最後のステップに入力・出力の指定がなければ、レスポンスから補う（描画したテンプレートと画面のタイトル）
     *
     * @param  array<int, array<string, mixed>>  $steps
     * @return array<int, array<string, mixed>>
     */
    private function completeLastStep(array $steps, Response $response): array
    {
        $endNo = array_key_last($steps);
        if ($endNo === null || $response->getStatusCode() !== 200) {
            return $steps;
        }

        if (! $steps[$endNo]['in']) {
            // 画面描画の入力情報を補完
            $original = $response instanceof \Illuminate\Http\Response ? $response->getOriginalContent() : null;
            if (is_object($original) && method_exists($original, 'getPath')) {
                $viewPath = str_replace(resource_path('views/'), '', $original->getPath());
                $viewPath && $steps[$endNo]['in'] = ['テンプレート' => '!'.$viewPath];
            }
        }

        if (! $steps[$endNo]['out']) {
            // 画面描画の出力情報を補完
            $contents = $response->getContent();
            if (is_string($contents) && $contents !== '') {
                $contents = str_replace(["\r\n", "\r", "\n"], '', substr($contents, 0, 200000));
                if (preg_match('/<title>(.*?)<\/title>/i', $contents, $match)) {
                    $steps[$endNo]['out'] = ['HTML' => '!'.trim(html_entity_decode($match[1], ENT_QUOTES | ENT_HTML5, 'UTF-8'))];
                }
            }
        }

        return $steps;
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

            // 配列の中の空の要素は空行として残す（配列でない空の指定は行を取らない）
            $keepsEmptyLine = is_array($items);
            foreach ($keepsEmptyLine ? $items : [$items] as $item) {
                $item = $this->stringify($item);
                if ($item !== '') {
                    $excel->set(1, $valueColumn, $line, $this->normalizeInOutValue($item));
                }

                ($item !== '' || $keepsEmptyLine) && $line++;
            }
        }

        return ++$line;
    }

    private function fillProcessRows($line, $step, $excel, &$headlineNo, &$indentNo)
    {
        // 見出しは K 列、処理は L 列。コメントと処理の箇所は、その1つ右の列に書く
        $detailColumn = 'M';

        $lastProcessKey = array_key_last($step['process']);

        foreach ($step['process'] as $processKey => $process) {
            $comments = preg_split('/\r\n|\r|\n/', $this->stringify($process));
            $process = array_shift($comments);

            // 行頭が#のときは見出し扱い
            $headline = str_starts_with($process, '#');
            $headline && $process = trim(substr($process, 1));
            $process = $this->normalizeProcessValue($process);

            // 空の処理（null や空文字）は、番号だけの行を出さず、連番も進めない
            if (trim($process) !== '') {
                if ($headline) {
                    $excel->set(1, 'K', $line, $headlineNo.'. '.$process);
                    $headlineNo++;
                    $indentNo = 1;
                } else {
                    $excel->set(1, 'L', $line, $indentNo.') '.$process);
                    $indentNo++;
                }
                $detailColumn = $headline ? 'L' : 'M';
                $line++;
            }

            // 追加コメントを補完
            $addComment = $this->stringify($this->findSupplementaryComment($process));
            $addComment === '' || $comments = array_merge($comments, preg_split('/\r\n|\r|\n/', $addComment));

            // バリデーション情報は、ステップの最後の処理に一度だけ付ける
            if ($processKey === $lastProcessKey) {
                $comments = array_merge($comments, $this->formatValidationComments($step['validate']));
            }

            foreach ($comments as $comment) {
                $excel->set(1, $detailColumn, $line, $this->normalizeProcessValue($comment));
                $line++;
            }
        }

        // 処理の箇所を記述
        $path = str_replace(base_path().'/', '', $step['path']);
        $excel->set(1, $detailColumn, $line, $path.'@'.$step['function'].':'.$step['line']);
        $line++;

        return ++$line;
    }

    /**
     * バリデーション情報をコメント行に整形する
     *
     * @param  array<int|string, mixed>  $validates  ['name' => 入力値, 'validate' => 条件, 'message' => メッセージ] の配列。
     *                                               「入力値 => 条件」や、入力値の名前だけ（文字列、または 名前 => []）でも受け付ける
     * @return array<int, string>
     */
    private function formatValidationComments(array $validates): array
    {
        $comments = [];

        // ['name' => 入力値, 'validate' => 条件] を1件だけ、配列で包まずに渡した場合
        // （「入力値 => 条件」の書き方で name という入力値がある場合と取り違えないよう、
        //   キーが name / validate / message だけで、validate か message を含むときに限る）
        $detailKeys = ['name' => 1, 'validate' => 1, 'message' => 1];
        if (isset($validates['name']) && ! is_array($validates['name'])
            && ! array_diff_key($validates, $detailKeys)
            && (isset($validates['validate']) || isset($validates['message']))) {
            $validates = [$validates];
        }

        foreach ($validates as $key => $validate) {
            $hasDetailKeys = is_array($validate) && array_intersect_key($validate, $detailKeys);

            if ($hasDetailKeys) {
                // ['name' => 入力値, 'validate' => 条件, 'message' => メッセージ]
                $name = $this->stringify($validate['name'] ?? (is_string($key) ? $key : ''));
                $rule = $this->stringify($validate['validate'] ?? '');
                $message = $this->stringify($validate['message'] ?? '');
            } elseif (is_string($key)) {
                // 入力値 => 条件（Laravel のルールと同じ書き方。条件は文字列でも配列でもよい）
                $name = $key;
                $rule = $this->stringify($validate);
                $message = '';
            } else {
                // 入力値の名前だけ
                $name = $this->stringify($validate);
                $rule = '';
                $message = '';
            }

            if ($name === '') {
                continue;
            }

            $comment = '・'.$name;
            $rule === '' || $comment .= ': '.$rule;
            $message === '' || $comment .= ': '.$message;
            $comments[] = $comment;
        }

        // 見出しは、表示する項目があるときだけ付ける
        $comments && array_unshift($comments, '<入力値>: <条件>: <メッセージ>');

        return $comments;
    }

    /**
     * docs() に渡された値を文字列にする（配列は要素を「, 」でつなぐ）
     */
    private function stringify($value): string
    {
        if (is_array($value)) {
            return implode(', ', array_map([$this, 'stringify'], $value));
        }

        if (is_scalar($value) || $value instanceof \Stringable) {
            return (string) $value;
        }

        return '';
    }

    /**
     * 設定ファイル（docs/common.php と docs/コントローラー名.php）を読み込み、画面の概要と置き換えの設定を反映する
     * 後から読む設定ほど優先する（共通 → コントローラー → メソッド）
     *
     * 設定ファイルはこのメソッドの中で読み込む。設定ファイルからは $routeClass / $routeMethod / $excel と、
     * 読み込み済みの $config / $keywords / $neglectPatterns / $commentMap を参照・変更できる
     */
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

            foreach ([$config, $config[$routeMethod] ?? []] as $level) {
                // キーワードと追加コメントは、同じキーなら後の設定で上書きする
                // （array_merge は '200' のような数字のキーを 0, 1, … に振り直してしまうため使わない）
                isset($level['keyword']) && $keywords = array_replace($keywords, $level['keyword']);
                isset($level['comment']) && $commentMap = array_replace($commentMap, $level['comment']);

                // 非表示行は追加していく
                isset($level['neglect']) && $neglectPatterns = array_merge($neglectPatterns, $level['neglect']);
            }
        }

        // 入出力のキーワードは長いものを優先して照合するので、長い順（同じ長さなら昇順）に並べておく
        $this->keywords = $keywords;
        $this->keywordsByLength = array_keys($keywords);
        usort($this->keywordsByLength, fn ($a, $b) => [strlen((string) $b), (string) $a] <=> [strlen((string) $a), (string) $b]);

        // コントローラーの設定が共通設定に書き足す形だと、共通の非表示行が二重に入るので重複を除く
        $this->neglectPatterns = array_values(array_unique($neglectPatterns, SORT_REGULAR));
        $this->commentMap = $commentMap;
    }

    private function normalizeInOutValue($item)
    {
        if (str_starts_with($item, '!')) {
            return substr($item, 1);
        }

        foreach ($this->keywordsByLength as $key) {
            $key = (string) $key;
            if ($key !== '' && strpos($item, $key) !== false) {
                // キーワード置き換え
                $item = str_replace($key, $key.': '.$this->stringify($this->keywords[$key]), $item);

                return $item;
            }
        }

        return $item;
    }

    private function normalizeProcessValue($item)
    {
        if (str_starts_with($item, '!')) {
            return substr($item, 1);
        }

        foreach ($this->keywords as $key => $value) {
            // キーワード置き換え
            $item = str_replace('<'.$key.'>', '<'.$this->stringify($value).'>', $item);
        }

        return $item;
    }

    private function shouldSkipStep($item)
    {
        $item = $this->stripWhitespace($item);
        foreach ($this->neglectPatterns as $neglect) {
            // 空の指定は、すべてのステップに一致してしまうので無視する
            $neglect = $this->stripWhitespace($this->stringify($neglect));
            if ($neglect !== '' && strpos($item, $neglect) !== false) {
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
        // 名前空間の App\Http\Controllers は省く。途中にある場合（Modules\Blog\App\Http\Controllers）や
        // 名前の一部が一致する場合（MyApp\Http\Controllers → My）も、従来どおり同じ規則で省く
        // （規則を変えると、既存の設計書と設定ファイルの場所が変わってしまう）
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
