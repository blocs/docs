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
}
