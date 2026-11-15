<?php

namespace Blocs\Tests;

use Blocs\Excel;
use Blocs\Middleware\Docs;
use Illuminate\Container\Container;
use Illuminate\Contracts\Debug\ExceptionHandler;
use Illuminate\Events\Dispatcher;
use Illuminate\Http\Request;
use Illuminate\Http\Response;
use Illuminate\Routing\CallableDispatcher;
use Illuminate\Routing\Router;
use Illuminate\Support\Facades\Facade;
use PHPUnit\Framework\TestCase;
use ReflectionMethod;

class DocsMiddlewareTest extends TestCase
{
    public function test_normalize_route_action_uses_invoke_for_invokable_controllers(): void
    {
        $this->assertSame(
            'Admin/UserController@__invoke',
            Docs::normalizeRouteAction('App\\Http\\Controllers\\Admin\\UserController')
        );
    }

    public function test_normalize_route_action_keeps_controller_method(): void
    {
        $this->assertSame(
            'Admin/UserController@index',
            Docs::normalizeRouteAction('App\\Http\\Controllers\\Admin\\UserController@index')
        );
    }

    public function test_normalize_route_action_falls_back_when_action_is_missing(): void
    {
        $this->assertSame('class@method', Docs::normalizeRouteAction(null));
        $this->assertSame('class@method', Docs::normalizeRouteAction(''));
    }

    public function test_supplementary_comment_matches_keys_that_contain_spaces(): void
    {
        $docs = new Docs;
        $commentMap = (new \ReflectionClass($docs))->getProperty('commentMap');
        $commentMap->setValue($docs, ['hello world' => 'the comment']);

        $method = new ReflectionMethod(Docs::class, 'findSupplementaryComment');

        $this->assertSame('the comment', $method->invoke($docs, 'prefix hello world suffix'));
    }

    public function test_validation_rows_are_written_once_per_step(): void
    {
        $docs = new Docs;
        $excel = $this->recordingExcel();
        $headlineNo = 1;
        $indentNo = 1;

        $step = [
            'path' => '/app/Controller.php',
            'function' => 'store',
            'line' => 10,
            'process' => ['#登録する', '値を検証する', 'DBに保存する'],
            'validate' => [['name' => 'email', 'validate' => 'required', 'message' => '必須です']],
        ];

        // 処理の箇所を書き出すときに base_path() を使うため、最小限のアプリケーションを用意する
        $previousContainer = Container::getInstance();
        Container::setInstance(new class extends Container
        {
            public function basePath($path = '')
            {
                return '/app'.($path === '' ? '' : '/'.$path);
            }
        });

        try {
            $method = new ReflectionMethod(Docs::class, 'fillProcessRows');
            $method->invokeArgs($docs, [5, $step, $excel, &$headlineNo, &$indentNo]);
        } finally {
            Container::setInstance($previousContainer);
        }

        // バリデーションの一覧は、ステップの最後の処理に一度だけ出力する
        $written = implode("\n", $excel->cells);
        $this->assertSame(1, substr_count($written, '<入力値>: <条件>: <メッセージ>'));
        $this->assertSame(1, substr_count($written, '・email: required: 必須です'));
        $this->assertStringContainsString("L8=2) DBに保存する\nM9=<入力値>: <条件>: <メッセージ>\nM10=・email: required: 必須です\nM11=Controller.php@store:10", $written);
    }

    public function test_validation_accepts_names_without_details(): void
    {
        $docs = new Docs;
        $method = new ReflectionMethod(Docs::class, 'formatValidationComments');

        // docs() に名前だけ渡すと ['email' => []]、名前のリストなら ['email', 'name'] になる
        $this->assertSame(['<入力値>: <条件>: <メッセージ>', '・email'], $method->invoke($docs, ['email' => []]));
        $this->assertSame(['<入力値>: <条件>: <メッセージ>', '・email', '・name'], $method->invoke($docs, ['email', 'name']));
        $this->assertSame(
            ['<入力値>: <条件>: <メッセージ>', '・email: required, max:5'],
            $method->invoke($docs, [['name' => 'email', 'validate' => ['required', 'max:5']]])
        );
        $this->assertSame([], $method->invoke($docs, []));
    }

    public function test_non_string_values_do_not_break_rows(): void
    {
        $docs = new Docs;
        $excel = $this->recordingExcel();

        $method = new ReflectionMethod(Docs::class, 'fillIoRows');
        $method->invoke($docs, 5, ['POST' => ['name', 5, ['a', 'b'], null, '']], $excel, ['A' => null], 'B');

        $this->assertSame(['A6=POST', 'B7=name', 'B8=5', 'B9=a, b'], $excel->cells);
    }

    public function test_empty_neglect_pattern_does_not_hide_every_step(): void
    {
        $docs = new Docs;
        $neglectPatterns = (new \ReflectionClass($docs))->getProperty('neglectPatterns');
        $neglectPatterns->setValue($docs, ['', '  ', '非表示 の 処理']);

        $method = new ReflectionMethod(Docs::class, 'shouldSkipStep');

        // 空の指定はどのステップにも一致させない
        $this->assertFalse($method->invoke($docs, '通常の処理'));
        $this->assertTrue($method->invoke($docs, 'これは非表示の処理です'));
    }

    public function test_docs_helper_normalizes_arguments(): void
    {
        $GLOBALS['DOC_GENERATOR'] = [];

        try {
            docs('処理だけ');
            docs('入力', '処理');
            docs(['POST' => ['name']], ['# 見出し', '処理'], '出力', 'email');
            docs(null, '処理', null, [['name' => 'email']]);
            docs(new \stdClass, '処理', 1.5);

            $steps = $GLOBALS['DOC_GENERATOR'];
        } finally {
            unset($GLOBALS['DOC_GENERATOR']);
        }

        $pick = fn (array $step) => [$step['in'], $step['process'], $step['out'], $step['validate']];

        $this->assertSame([[], ['処理だけ'], [], []], $pick($steps[0]));
        $this->assertSame([['入力' => []], ['処理'], [], []], $pick($steps[1]));
        $this->assertSame([['POST' => ['name']], ['# 見出し', '処理'], ['出力' => []], ['email' => []]], $pick($steps[2]));
        $this->assertSame([[], ['処理'], [], [['name' => 'email']]], $pick($steps[3]));

        // 配列のキーにできない値は記録せず、例外にもしない
        $this->assertSame([[], ['処理'], ['1.5' => []], []], $pick($steps[4]));

        $this->assertSame(__FUNCTION__, $steps[0]['function']);
        $this->assertSame(__FILE__, $steps[0]['path']);

        // 生成が無効なときは何も記録しない
        docs('記録されない');
        $this->assertArrayNotHasKey('DOC_GENERATOR', $GLOBALS);
    }

    public function test_docs_helper_accepts_named_arguments(): void
    {
        $GLOBALS['DOC_GENERATOR'] = [];

        try {
            // 名前付き引数や文字列キーの配列の展開でも、警告を出さずに順番どおり扱う
            docs(in: 'a', process: 'p', out: 'o');
            docs(...['x' => 'a', 'y' => 'p']);
            $steps = $GLOBALS['DOC_GENERATOR'];
        } finally {
            unset($GLOBALS['DOC_GENERATOR']);
        }

        $this->assertSame([['a' => []], ['p'], ['o' => []]], [$steps[0]['in'], $steps[0]['process'], $steps[0]['out']]);
        $this->assertSame([['a' => []], ['p']], [$steps[1]['in'], $steps[1]['process']]);
    }

    public function test_config_files_keep_numeric_keywords_and_see_route_variables(): void
    {
        $basePath = sys_get_temp_dir().'/blocs-docs-config-'.bin2hex(random_bytes(4));
        mkdir($basePath.'/docs/Admin', 0755, true);
        file_put_contents($basePath.'/docs/common.php', '<?php $config = ["keyword" => ["200" => "OK", "404" => "NotFound", "name" => "名前"], "neglect" => ["共通で隠す"], "comment" => ["a" => "共通"]];');
        file_put_contents($basePath.'/docs/Admin/UserController.php', '<?php $config["description"] = "class=".$routeClass." method=".$routeMethod; $config["index"] = ["keyword" => ["404" => "見つからない"], "neglect" => ["メソッドで隠す"], "comment" => ["a" => "メソッド"]];');

        $previousContainer = Container::getInstance();
        Container::setInstance(new class($basePath) extends Container
        {
            public function __construct(private string $root) {}

            public function basePath($path = '')
            {
                return $this->root.($path === '' ? '' : '/'.$path);
            }
        });

        try {
            $docs = new Docs;
            $excel = $this->recordingExcel();
            (new ReflectionMethod(Docs::class, 'loadConfig'))->invoke($docs, 'Admin/UserController', 'index', $excel);

            $property = fn (string $name) => (new \ReflectionClass($docs))->getProperty($name)->getValue($docs);

            // 数字のキーワードが 0, 1 に振り直されない。メソッドの設定が共通の設定より優先される
            $this->assertSame(['200' => 'OK', '404' => '見つからない', 'name' => '名前'], $property('keywords'));
            $this->assertSame(['共通で隠す', 'メソッドで隠す'], $property('neglectPatterns'));
            $this->assertSame(['a' => 'メソッド'], $property('commentMap'));

            // 設定ファイルから $routeClass / $routeMethod を参照でき、共通設定への書き足しもできる
            $this->assertContains('Z1=class=Admin/UserController method=index', $excel->cells);

            $normalize = new ReflectionMethod(Docs::class, 'normalizeInOutValue');
            $this->assertSame('200: OK', $normalize->invoke($docs, '200'));
            $this->assertSame('404: 見つからない', $normalize->invoke($docs, '404'));
        } finally {
            Container::setInstance($previousContainer);
            array_map('unlink', [$basePath.'/docs/common.php', $basePath.'/docs/Admin/UserController.php']);
            rmdir($basePath.'/docs/Admin');
            rmdir($basePath.'/docs');
            rmdir($basePath);
        }
    }

    public function test_validation_accepts_laravel_style_rules(): void
    {
        $docs = new Docs;
        $method = new ReflectionMethod(Docs::class, 'formatValidationComments');
        $heading = '<入力値>: <条件>: <メッセージ>';

        // 入力値 => 条件（文字列・配列）
        $this->assertSame([$heading, '・email: required', '・pw: required, min:8'], $method->invoke($docs, ['email' => 'required', 'pw' => ['required', 'min:8']]));

        // 1件だけを配列で包まずに渡した場合
        $this->assertSame([$heading, '・email: required: 必須です'], $method->invoke($docs, ['name' => 'email', 'validate' => 'required', 'message' => '必須です']));

        // name という入力値がある「入力値 => 条件」を、1件だけの指定と取り違えない
        $this->assertSame([$heading, '・name: required', '・email: required|email'], $method->invoke($docs, ['name' => 'required', 'email' => 'required|email']));
        $this->assertSame([$heading, '・email: required', '・name: required|max:10'], $method->invoke($docs, ['email' => 'required', 'name' => 'required|max:10']));
    }

    public function test_empty_process_and_crlf_do_not_leave_artifacts(): void
    {
        $docs = new Docs;
        $excel = $this->recordingExcel();
        $headlineNo = 1;
        $indentNo = 1;

        $step = ['path' => '/app/C.php', 'function' => 'f', 'line' => 1, 'process' => [null, "処理\r\nコメント", ''], 'validate' => []];

        $previousContainer = Container::getInstance();
        Container::setInstance(new class extends Container
        {
            public function basePath($path = '')
            {
                return '/app'.($path === '' ? '' : '/'.$path);
            }
        });

        try {
            (new ReflectionMethod(Docs::class, 'fillProcessRows'))->invokeArgs($docs, [5, $step, $excel, &$headlineNo, &$indentNo]);
        } finally {
            Container::setInstance($previousContainer);
        }

        // 空の処理は番号だけの行を出さず連番も進めない。コメントに \r を残さない
        $this->assertSame(['L6=1) 処理', 'M7=コメント', 'M8=C.php@f:1'], $excel->cells);
        $this->assertSame(2, $indentNo);
    }

    public function test_empty_items_inside_array_keep_their_line(): void
    {
        $docs = new Docs;
        $excel = $this->recordingExcel();

        // 配列の中の空の要素は空行として残す（従来どおりの位置に出力する）。配列でない空の指定は行を取らない
        $nextLine = (new ReflectionMethod(Docs::class, 'fillIoRows'))->invoke($docs, 5, ['DB' => ['users', '', 'b'], 'X' => ''], $excel, ['A' => null], 'B');

        $this->assertSame(['A6=DB', 'B7=users', 'B9=b', 'A10=X'], $excel->cells);
        $this->assertSame(11, $nextLine);
    }

    public function test_normalize_route_action_keeps_the_same_rule_as_before(): void
    {
        $this->assertSame('Admin/UserController@index', Docs::normalizeRouteAction('App\\Http\\Controllers\\Admin\\UserController@index'));

        // 途中や名前の一部にある場合も、v2.0.16 と同じ規則で省く
        // （規則を変えると、既存の設計書と設定ファイルの場所が変わってしまう）
        $this->assertSame('Modules/Blog/PostController@index', Docs::normalizeRouteAction('Modules\\Blog\\App\\Http\\Controllers\\PostController@index'));
        $this->assertSame('Modules/MyX@index', Docs::normalizeRouteAction('Modules\\MyApp\\Http\\Controllers\\X@index'));
    }

    public function test_handle_generates_document_without_affecting_response(): void
    {
        // アプリのパスに空白と日本語を含める
        $basePath = sys_get_temp_dir().'/blocs docs アプリ-'.bin2hex(random_bytes(4));
        mkdir($basePath.'/docs', 0755, true);
        copy(dirname(__DIR__, 2).'/docs/format.xlsx', $basePath.'/docs/format.xlsx');

        $previousContainer = Container::getInstance();
        $previousFacadeApplication = Facade::getFacadeApplication();

        $app = new class($basePath) extends Container
        {
            public bool $production = false;

            public function __construct(private string $root) {}

            public function basePath($path = '')
            {
                return $this->root.($path === '' ? '' : '/'.$path);
            }

            public function resourcePath($path = '')
            {
                return $this->root.'/resources'.($path === '' ? '' : '/'.$path);
            }

            public function isProduction()
            {
                return $this->production;
            }
        };
        $handler = new class implements ExceptionHandler
        {
            /** @var array<int, string> */
            public array $reported = [];

            public function report(\Throwable $e)
            {
                $this->reported[] = $e->getMessage();
            }

            public function shouldReport(\Throwable $e)
            {
                return true;
            }

            public function render($request, \Throwable $e)
            {
                return new Response('error', 500);
            }

            public function renderForConsole($output, \Throwable $e) {}
        };

        Container::setInstance($app);
        Facade::clearResolvedInstances();
        Facade::setFacadeApplication($app);
        $router = new Router(new Dispatcher($app), $app);
        $app->instance('router', $router);
        $app->instance(ExceptionHandler::class, $handler);
        $app->bind(\Illuminate\Routing\Contracts\CallableDispatcher::class, fn ($app) => new CallableDispatcher($app));

        $router->get('/screen', function () {
            docs(['GET' => ['id']], ['# 見出し', '値を検証する'], ['HTML' => '結果'], [['name' => 'id', 'validate' => 'required', 'message' => '必須です']]);
            docs('結果を返す');

            return new Response('<html><title>画面 &amp; タイトル</title></html>');
        })->middleware([Docs::class]);

        $documentPath = $basePath.'/docs/class@method.xlsx';
        $dispatch = fn () => $router->dispatch(Request::create('/screen', 'GET'));

        try {
            // 通常: レスポンスはそのまま返り、設計書が生成され、記録用のグローバル変数は片付く
            $response = $dispatch();
            $this->assertSame(200, $response->getStatusCode());
            $this->assertStringContainsString('<title>', $response->getContent());
            $this->assertSame([], $handler->reported);
            $this->assertArrayNotHasKey('DOC_GENERATOR', $GLOBALS);

            $excel = new Excel($documentPath);
            $this->assertSame(['method'], $excel->sheetNames());
            $this->assertSame('class@method', $excel->get(1, 'E', '2'));
            $this->assertSame('GET', $excel->get(1, 'A', '6'));
            $this->assertSame('1. 見出し', $excel->get(1, 'K', '6'));
            $this->assertSame('1) 値を検証する', $excel->get(1, 'L', '7'));
            $this->assertSame('<入力値>: <条件>: <メッセージ>', $excel->get(1, 'M', '8'));
            $this->assertSame('・id: required: 必須です', $excel->get(1, 'M', '9'));
            $this->assertSame('2) 結果を返す', $excel->get(1, 'L', '12'));
            $this->assertSame('画面 & タイトル', $excel->get(1, 'AQ', '13'));
            unset($excel);

            // ミドルウェアが二重に登録されていても（グローバルとルートの両方など）、設計書の内容は同じ
            $router->get('/twice', function () {
                docs('二重登録の処理');

                return new Response('ok');
            })->middleware([Docs::class, Docs::class.':again']);
            $router->dispatch(Request::create('/twice', 'GET'));
            $this->assertSame('1) 二重登録の処理', (new Excel($documentPath))->get(1, 'L', '6'));
            $this->assertSame([], $handler->reported);
            $this->assertArrayNotHasKey('DOC_GENERATOR', $GLOBALS);

            // 同じリクエストのオブジェクトをもう一度処理しても、設計書は更新される
            $reused = Request::create('/twice', 'GET');
            unlink($documentPath);
            $router->dispatch($reused);
            unlink($documentPath);
            $router->dispatch($reused);
            $this->assertFileExists($documentPath);

            // duplicate() で作ったリクエストは、元のリクエストとは別に扱う
            $duplicated = $reused->duplicate();
            $this->assertNull($duplicated->attributes->get(Docs::class));

            // 保存できないとき: レスポンスはそのまま返り、失敗が記録される
            // （root はパーミッションに関係なく書き込めるので、Docker などでは確かめられない）
            unlink($documentPath);
            chmod($basePath.'/docs', 0555);
            clearstatcache();
            if (! is_writable($basePath.'/docs')) {
                $response = $dispatch();
                $this->assertSame(200, $response->getStatusCode());
                $this->assertCount(1, $handler->reported);
                $this->assertStringContainsString('Can not save document', $handler->reported[0]);
                $this->assertFileDoesNotExist($documentPath);
                $this->assertArrayNotHasKey('DOC_GENERATOR', $GLOBALS);
            }
            chmod($basePath.'/docs', 0755);

            // 本番環境: 生成しない
            $app->production = true;
            $this->assertSame(200, $dispatch()->getStatusCode());
            $this->assertFileDoesNotExist($documentPath);
        } finally {
            Container::setInstance($previousContainer);
            Facade::clearResolvedInstances();
            Facade::setFacadeApplication($previousFacadeApplication);

            is_dir($basePath.'/docs') && chmod($basePath.'/docs', 0755);
            foreach (array_diff(scandir($basePath.'/docs'), ['.', '..']) as $fileName) {
                unlink($basePath.'/docs/'.$fileName);
            }
            rmdir($basePath.'/docs');
            rmdir($basePath);
        }
    }

    private function recordingExcel(): object
    {
        return new class
        {
            /** @var array<int, string> */
            public array $cells = [];

            public function set($sheetNo, $column, $row, $value)
            {
                // 実際のシートと同じ番地で記録する（int の行は0始まり、文字列の行は1始まり）
                $this->cells[] = $column.(is_int($row) ? $row + 1 : $row).'='.$value;

                return $this;
            }
        };
    }
}
