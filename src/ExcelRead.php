<?php
/**
 * Created by PhpStorm.
 * User: xlt
 * Date: 2016/4/22
 * Time: 9:17.
 */

namespace xltxlm\phpexcel;

use PHPExcel_Shared_Date;

/**
 * 解析出excel的内容,返回数组,第一行是标题头
 * Class ExcelRead.
 */
final class ExcelRead
{
    /** @var string 需要解析的Excel本地文件 */
    protected $excelFile = '';

    private $excelObject;
    /** @var \PHPExcel_Worksheet */
    private $currentSheet;
    /** @var int 当前有多少行数据 */
    private $rowNum = 0;
    /** @var int 有多少列数据 */
    private $columnNum = 0;
    /** @var array 表格头数据 */
    private $firstRow = [];

    private $isFormatCell = true;

    private $emptyColumns = [];

    /**
     * @param string $excelFile
     *
     * @return ExcelRead
     */
    public function setExcelFile($excelFile)
    {
        $this->excelFile = $excelFile;

        $this->excelObject = (new \PHPExcel_Reader_Excel2007())->load($this->excelFile);
        $this->currentSheet = $this->excelObject->getSheet(0);
        $this->rowNum = $this->currentSheet->getHighestRow();

        $cols = 0;
        $strlen = strlen($this->currentSheet->getHighestColumn());
        for ($i = 1; $i <= $strlen; $i++) {
            $i1 = ord($this->currentSheet->getHighestColumn()[$i - 1]) - 64;
            $i2 = ($strlen - $i) * $i1 * 26 ?: $i1;
            $cols += $i2;
        }
        $this->columnNum = $cols;
        $this->firstRow = $this->findFirstRow();

        return $this;
    }

    /**
     * 获取Excel的行头.
     *
     * @throws \PHPExcel_Exception
     *
     * @return array
     */
    private function findFirstRow()
    {
        $firstRow = [];
        for ($key = 0; $key < $this->columnNum; ++$key) {
            //A-Z溢出
            $char = chr(65 + $key % 26);
            if ($key >= 26) {
                $char = chr(64 + floor($key / 26)).$char;
            }

            $cell = trim(
                $this->currentSheet->getCell($char. 1)
                    ->getValue()
            );
            $firstRow[] = $cell;
            if (empty($cell)) {
                $this->emptyColumns[] = $key;
            }
        }

        return $firstRow;
    }

    /*检测每行数据是否为空*/

    /**
     * @return array
     */
    public function getFirstRow()
    {
        return $this->firstRow;
    }

    /**
     * 是否格式化Excel中的时间数据(会将小于1的数字格式化为时间, 再输出)
     * 默认格式化.
     *
     * @param bool $isFormatCell
     *
     * @return ExcelRead
     */
    public function setIsFormatTimeCell($isFormatCell)
    {
        $this->isFormatCell = $isFormatCell;

        return $this;
    }

    /**
     * @throws \PHPExcel_Exception
     *
     * @return array
     */
    public function __invoke()
    {
        $excelData = [];
        $i = 0;
        for ($currentRow = 2; $currentRow <= $this->rowNum; ++$currentRow) {
            for ($key = 0; $key < $this->columnNum; ++$key) {
                if (in_array($key, $this->emptyColumns)) {
                    continue;
                }
                //A-Z溢出
                $char = chr(65 + $key % 26);
                if ($key >= 26) {
                    $char = chr(64 + floor($key / 26)).$char;
                }
                $cell = $this->currentSheet->getCell($char.$currentRow);
                $cellData = preg_replace('/^[　]*/', '', $cell->getValue());    //去除开头的中文空格
                $cellData = trim($cellData);

                //把0.xxxx的数字转换成时间
                $isTime = preg_match('/^0\.(d+)?/', $cellData);
                if ($this->isFormatCell && $isTime) {
                    $cellData = \PHPExcel_Style_NumberFormat::toFormattedString($cellData, 'hh:mm');
                }
                //把40000~60000的数字转换成日期
                $isDate = is_numeric($cellData) && $cellData > 40000 && $cellData < 60000;
                if ($this->isFormatCell && $isDate) {
                    //$cellData 含有小数
                    if (strpos($cellData, '.')) {
                        $cellData = gmdate('Y-m-d H:i:s', PHPExcel_Shared_Date::ExcelToPHP($cellData));
                    } else {
                        $cellData = \PHPExcel_Style_NumberFormat::toFormattedString(
                            $cellData,
                            \PHPExcel_Style_NumberFormat::FORMAT_DATE_YYYYMMDD2
                        );
                    }
                }

                $excelData[$i][$this->firstRow[$key]] = $cellData;
            }

            if ($this->checkIsEmpty($excelData[$i])) {
                unset($excelData[$i]);
                $i ?: $i--;
            }
            $i++;
        }

        return $excelData;
    }

    private function checkIsEmpty($inputData)
    {
        $flag = true;
        foreach ($inputData as $value) {
            if (!empty($value)) {
                $flag = false;
                break;
            }
        }

        return $flag;
    }
}
