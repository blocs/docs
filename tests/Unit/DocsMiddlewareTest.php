<?php

namespace Blocs\Tests;

use Blocs\Middleware\Docs;
use Illuminate\Container\Container;
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
        $this->assertStringContainsString("L7=2) DBに保存する\nM8=<入力値>: <条件>: <メッセージ>\nM9=・email: required: 必須です\nM10=Controller.php@store:10", $written);
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

        $this->assertSame(['A5=POST', 'B6=name', 'B7=5', 'B8=a, b'], $excel->cells);
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

    private function recordingExcel(): object
    {
        return new class
        {
            /** @var array<int, string> */
            public array $cells = [];

            public function set($sheetNo, $column, $row, $value)
            {
                $this->cells[] = $column.$row.'='.$value;

                return $this;
            }
        };
    }
}
