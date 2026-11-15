<div align="center"><img src="logo.svg" width="400" /></div>

# Laravel Document Generator
Laravelのドキュメント生成ツール

[![Latest stable version](https://img.shields.io/packagist/v/blocs/docs)](https://packagist.org/packages/blocs/docs)
[![Total downloads](https://img.shields.io/packagist/dt/blocs/docs)](https://packagist.org/packages/blocs/docs)
[![GitHub code size](https://img.shields.io/github/languages/code-size/blocs/docs)](https://github.com/blocs/docs)
[![GitHub license](https://img.shields.io/github/license/blocs/docs)](https://github.com/blocs/docs)
[![Laravel awesome](https://img.shields.io/badge/Awesome-Laravel-green)](https://github.com/blocs/docs)
[![Laravel version](https://img.shields.io/badge/laravel-%3E%3D10-green)](https://github.com/blocs/docs)
[![PHP version](https://img.shields.io/badge/php-%3E%3D8.3-blue)](https://github.com/blocs/docs)

[**Website**](https://blocs.jp/)
| [**iknow**](https://linear.iknow.dev/?category=2)

## 導入方法
1. PHP と Composer をセットアップ
2. Composer で Laravel をインストール
```sh
composer require blocs/docs
```
3. docsをコピー
```sh
cp -r vendor/blocs/docs/docs ./
```

## ドキュメントの生成
コントローラーの処理を `docs()` で記録し、画面ごとの設計書（Excel）を `docs/` に書き出します。

1. ミドルウェアを登録（`bootstrap/app.php`）
```php
->withMiddleware(function (Middleware $middleware) {
    $middleware->web(append: [\Blocs\Middleware\Docs::class]);
})
```
2. コントローラーで処理を記録
```php
// 処理だけ
docs('# 一覧画面の表示');

// 入力, 処理
docs(['POST' => '入力値'], '入力値を検証する');

// 入力, 処理, 出力, バリデーション
docs(['POST' => '入力値'], 'データを追加する', ['データベース' => 'users'], [
    ['name' => 'email', 'validate' => 'required', 'message' => '必須です'],
]);
```
3. 画面を表示すると `docs/コントローラー名@メソッド名.xlsx` が生成されます

- `docs/format.xlsx`（雛形）がない場合と、本番環境（`APP_ENV=production`）では生成しません
- 生成に失敗しても画面の表示には影響しません（例外はログに記録されます）
- 画面名やキーワードの置き換えは `docs/common.php` と `docs/コントローラー名.php` で設定します

## Excel の読み書き
```php
use Blocs\Excel;

$excel = new Excel('template.xlsx');

// 読み取り（シートは番号（1始まり）または名前、列は番号（0始まり）または名前、行は int なら0始まり・文字列なら1始まり）
$value = $excel->get(1, 'A', '1');
$rows = $excel->all('Sheet1');

// 大きなファイルは1行ずつ読む
$excel->open(1);
while (($row = $excel->next()) !== false) {
    // ...
}

// 書き込み
$excel->set(1, 'A', '1', '文字列')
    ->set(1, 'B', '1', 123)
    ->set(1, 'C', '1', new DateTime('2024-01-02'))  // シリアル値で書き込み（セルに日付の書式が必要）
    ->name(1, '新しいシート名');

$excel->save('output.xlsx');      // 引数を省略すると読み込んだファイルを上書き
return $excel->download('出力.xlsx');
```
