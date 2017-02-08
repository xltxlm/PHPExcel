<?php
/**
 * Created by PhpStorm.
 * User: xialintai
 * Date: 2017/2/8
 * Time: 17:19
 */

namespace xltxlm\phpexcel\tests;


use PHPUnit\Framework\TestCase;
use xltxlm\phpexcel\ExcelRead;

class ExcelReadTest extends TestCase
{

    public function test26()
    {
        $excelRead = new ExcelRead();
        $datas = $excelRead
            ->setExcelFile(__DIR__.'/test.xlsx')
            ->__invoke();
        $this->assertEquals(101, count($excelRead->getFirstRow()));
        $this->assertEquals(10, count($datas));
    }
    public function test()
    {
        $excelRead = new ExcelRead();
        $datas = $excelRead
            ->setExcelFile(__DIR__.'/test2.xlsx')
            ->__invoke();
        $this->assertEquals(11, count($excelRead->getFirstRow()));
        $this->assertEquals(8, count($datas));
    }
}