<?php

namespace xltxlm\phpexcel;

/**
 * 生成excel文档 , 最大支持到 zz 列
 * Class make.
 */
final class ExcelWrite
{
    /** @var array 列数组 */
    protected $heads = [];
    /** @var array 全部数据集合 */
    protected $data = [];
    /** @var array 每一行的数据 */
    protected $dataRepeat = [];
    /** @var string 输出的文件名称 */
    protected $filename = '';
    /** @var bool 是否保存成文件 */
    protected $save2File = false;

    /**
     * @return string
     */
    public function getFilename(): string
    {
        return $this->filename;
    }

    /**
     * @param string $filename
     *
     * @return ExcelWrite
     */
    public function setFilename(string $filename): ExcelWrite
    {
        $this->filename = $filename;

        return $this;
    }

    /**
     * @return bool
     */
    public function isSave2File(): bool
    {
        return $this->save2File;
    }

    /**
     * @param bool $save2File
     *
     * @return ExcelWrite
     */
    public function setSave2File(bool $save2File): ExcelWrite
    {
        $this->save2File = $save2File;

        return $this;
    }

    /**
     * @return array
     */
    public function getData(): array
    {
        return $this->data;
    }

    /**
     * 如果是一个模型的类数组,那么需要转换一下,并且去掉索引,变成数字索引
     * @param array $data
     *
     * @return ExcelWrite
     */
    public function setData(array $data): ExcelWrite
    {
        $var = current($data);
        if (is_object($var) && method_exists($var, '__toArray')) {
            foreach ($data as &$datum) {
                eval('$datum = array_values($datum->__toArray());');
            }
        }
        $this->data = $data;

        return $this;
    }

    /**
     * @param mixed $dataRepeat
     *
     * @return ExcelWrite
     */
    public function setDataRepeat($dataRepeat): ExcelWrite
    {
        $this->data[] = $dataRepeat;

        return $this;
    }

    /**
     * @return array
     */
    public function getHeads(): array
    {
        return $this->heads;
    }

    /**
     * @param array $heads
     *
     * @return ExcelWrite
     */
    public function setHeads(array $heads): ExcelWrite
    {
        $this->heads = $heads;

        return $this;
    }

    /**
     * @desc 输出Excel
     *
     * @throws \PHPExcel_Exception
     * @throws \PHPExcel_Reader_Exception
     */
    public function __invoke()
    {
        $objPHPExcel = new \PHPExcel();
        $PHPExcel_Worksheet = $objPHPExcel->setActiveSheetIndex(0);

        //标记具体的合并单元格名称
        $colspanCol = [];

        //默认列头head开始插入的行号
        $coli = 1;
        //设置表格每行每列的值
        foreach ($this->getHeads() as $key => $v) {
            //A-Z溢出
            $char = chr(65 + $key % 26);
            if ($key >= 26) {
                $char = chr(64 + floor($key / 26)).$char;
            }
            $colspanCol[] = $char;

            //指定列宽

            $charlen = strlen($v) + 5;
            $objPHPExcel->getActiveSheet()->getColumnDimension($char)->setWidth($charlen);
            //设置列头值
            $PHPExcel_Worksheet->setCellValue($char.$coli, $v);

            //设置列标题内容居中
            $objPHPExcel->getActiveSheet()->getStyle($char.$coli)->getAlignment()->setHorizontal(
                \PHPExcel_Style_Alignment::HORIZONTAL_CENTER
            );
            $objPHPExcel->getActiveSheet()->getStyle($char.$coli)->getAlignment()->setVertical(
                \PHPExcel_Style_Alignment::VERTICAL_CENTER
            );

            //设置整列内容文字居中
            $objPHPExcel->getActiveSheet()->getStyle($char)->getAlignment()->setHorizontal(
                \PHPExcel_Style_Alignment::HORIZONTAL_CENTER
            );
            $objPHPExcel->getActiveSheet()->getStyle($char)->getAlignment()->setVertical(
                \PHPExcel_Style_Alignment::VERTICAL_CENTER
            );
        }
        //生成excel表格数据
        foreach ($this->getData() as $getalldata) {
            ++$coli;

            foreach ($getalldata as $key => $v) {
                //A-Z溢出
                $char = chr(65 + $key % 26);
                if ($key >= 26) {
                    $char = chr(64 + floor($key / 26)).$char;
                }
                //如果是数字
                $exsist = preg_match('#^[0-9\.,]+$#', $v);
                //以字符串的方式设置值,设定特定格式内容信息
                if ($exsist) {
                    $v = strtr($v, ["," => ""]);
                    $PHPExcel_Worksheet->setCellValueExplicit("{$char}{$coli}", $v, \PHPExcel_Cell_DataType::TYPE_NUMERIC);
                } else {
                    $PHPExcel_Worksheet->setCellValueExplicit("{$char}{$coli}", $v, \PHPExcel_Cell_DataType::TYPE_STRING2);
                }
            }
        }

        // Save Excel 2007 file
        $objWriter = \PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');

        if (empty($this->filename)) {
            $this->filename = date('YmdHis').'.xlsx';
        }
        if ($this->isSave2File()) {
            $objWriter->save($this->filename);
        } else {
            //直接输出excel到浏览器
            header('Content-Type: application/force-download');
            header('Content-disposition:attachment;filename="'.urlencode(basename($this->filename)));
            $filename = basename($this->filename);
            $encoded_filename = urlencode($filename);
            $encoded_filename = str_replace('+', '%20', $encoded_filename);
            if (preg_match('/MSIE/', $_SERVER['HTTP_USER_AGENT'])) {
                header('Content-Disposition: attachment; filename="'.$encoded_filename.'"');
            } elseif (preg_match('/Firefox/', $_SERVER['HTTP_USER_AGENT'])) {
                header('Content-Disposition: attachment; filename*="utf8\'\''.$filename.'"');
            } else {
                header('Content-Disposition: attachment; filename="'.$filename.'"');
            }

            header('Content-Transfer-Encoding: binary');
            header('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT');
            header('Cache-Control: must-revalidate, post-check=0, pre-check=0');
            header('Pragma: no-cache');
            $objWriter->save('php://output');
        }
    }
}
