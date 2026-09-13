<?php

namespace Blocs\Tests;

use Blocs\Excel;

require_once __DIR__.'/../ExcelTestCase.php';

/**
 * Excel::open() / first() / close() のストリーム読み取りのテスト
 */
class ExcelStreamTest extends ExcelTestCase
{
    public function test_first_reads_rows_sequentially(): void
    {
        $path = $this->buildXlsx([
            'Sheet1' => self::row(1, self::sharedCell('A1', 0))
                .self::row(2, self::sharedCell('A2', 1).self::numCell('B2', '2')),
        ], ['one', 'two']);
        $excel = new Excel($path);

        $excel->open(1);

        $this->assertSame(['one'], $excel->first());
        $this->assertSame(['two', '2'], $excel->first());

        // 終端に達するとfalse
        $this->assertFalse($excel->first());

        // 終端後に再度呼んでもfalse
        $this->assertFalse($excel->first());
    }

    public function test_blank_rows_are_filled(): void
    {
        $path = $this->buildXlsx([
            'Sheet1' => self::row(1, self::sharedCell('A1', 0))
                .self::row(4, self::sharedCell('A4', 1)),
        ], ['first', 'fourth']);
        $excel = new Excel($path);

        $excel->open(1);

        // 行2・行3は空配列で補完される
        $this->assertSame(['first'], $excel->first());
        $this->assertSame([], $excel->first());
        $this->assertSame([], $excel->first());
        $this->assertSame(['fourth'], $excel->first());
        $this->assertFalse($excel->first());
    }

    public function test_row_after_blank_rows_is_not_skipped(): void
    {
        $path = $this->buildXlsx([
            'Sheet1' => self::row(1, self::numCell('A1', '1'))
                .self::row(4, self::numCell('A4', '4'))
                .self::row(5, self::numCell('A5', '5')),
        ]);
        $excel = new Excel($path);

        $excel->open(1);

        $rows = [];
        while (($row = $excel->first()) !== false) {
            $rows[] = $row;
        }

        // 空白行の直後の行が読み飛ばされないこと（all()と同じ結果になる）
        $this->assertSame([['1'], [], [], ['4'], ['5']], $rows);
        $this->assertSame($rows, (new Excel($path))->all(1));
    }

    public function test_columns_filter(): void
    {
        $path = $this->buildXlsx([
            'Sheet1' => self::row(1, self::numCell('A1', '1').self::numCell('B1', '2').self::numCell('C1', '3')),
        ]);
        $excel = new Excel($path);

        $excel->open(1, [1]);

        $this->assertSame(['2'], $excel->first());
        $this->assertFalse($excel->first());
    }

    public function test_open_missing_sheet(): void
    {
        $path = $this->buildXlsx(['Sheet1' => '']);
        $excel = new Excel($path);

        $excel->open(9);

        $this->assertFalse($excel->first());
    }

    public function test_close_stops_reading(): void
    {
        $path = $this->buildXlsx([
            'Sheet1' => self::row(1, self::numCell('A1', '1'))
                .self::row(2, self::numCell('A2', '2')),
        ]);
        $excel = new Excel($path);

        $excel->open(1);
        $this->assertSame(['1'], $excel->first());

        $excel->close();

        $this->assertFalse($excel->first());
    }

    public function test_reopen_after_close(): void
    {
        $path = $this->buildXlsx([
            'Sheet1' => self::row(1, self::numCell('A1', '1')),
        ]);
        $excel = new Excel($path);

        $excel->open(1);
        $this->assertSame(['1'], $excel->first());
        $excel->close();

        // 再オープンで先頭から読み直せる
        $excel->open(1);
        $this->assertSame(['1'], $excel->first());
        $this->assertFalse($excel->first());
    }

    public function test_destructor_removes_temp_worksheet_when_path_contains_hash(): void
    {
        $dir = sys_get_temp_dir().DIRECTORY_SEPARATOR.'hash#'.uniqid('', true);
        mkdir($dir);
        $source = $this->buildXlsx([
            'Sheet1' => self::row(1, self::numCell('A1', '1')),
        ]);
        $path = $dir.DIRECTORY_SEPARATOR.'book.xlsx';
        $this->assertTrue(copy($source, $path));

        $compiled = (string) (config('view.compiled') ?? sys_get_temp_dir());
        $before = glob($compiled.DIRECTORY_SEPARATOR.'excel*') ?: [];

        $excel = new Excel($path);
        $excel->open(1);
        $this->assertSame(['1'], $excel->first());
        unset($excel);
        gc_collect_cycles();

        $after = glob($compiled.DIRECTORY_SEPARATOR.'excel*') ?: [];
        sort($before);
        sort($after);
        $this->assertSame($before, $after);

        is_file($path) && unlink($path);
        @rmdir($dir);
    }
}
