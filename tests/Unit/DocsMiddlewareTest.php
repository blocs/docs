<?php

namespace Blocs\Tests;

use Blocs\Middleware\Docs;
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
}
