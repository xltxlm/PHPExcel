<?php
/**
 * Created by PhpStorm.
 * User: xialintai
 * Date: 2017/2/8
 * Time: 16:56.
 */

namespace xltxlm\phpexcel\tests;

use PHPUnit\Framework\TestCase;
use xltxlm\phpexcel\ExcelWrite;

class ExcelWriteTest extends TestCase
{

    /**
     * 超出26列的数据
     */
    public function test26()
    {
        $header = [];
        for ($i = 0; $i <= 100; ++$i) {
            $header[] = 'id'.$i;
        }
        $data = [];
        for ($i = 0; $i < 10; ++$i) {
            $dataLine = [];
            for ($ii = 0; $ii <= 100; ++$ii) {
                $dataLine[] = $ii;
            }
            $data[] = $dataLine;
        }
        (new ExcelWrite())
            ->setHeads($header)
            ->setData($data)
            ->setFilename('test.xlsx')
            ->setSave2File(true)
            ->__invoke();
    }

    /**
     * 小于26列的数据
     */
    public function test()
    {
        $header = [];
        for ($i = 0; $i <= 10; ++$i) {
            $header[] = 'id'.$i;
        }
        $data = [];
        for ($i = 0; $i < 8; ++$i) {
            $dataLine = [];
            for ($ii = 0; $ii <= 10; ++$ii) {
                $dataLine[] = $ii;
            }
            $data[] = $dataLine;
        }
        (new ExcelWrite())
            ->setHeads($header)
            ->setData($data)
            ->setFilename('test2.xlsx')
            ->setSave2File(true)
            ->__invoke();
    }
}
