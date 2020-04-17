<?php

namespace PHPExcelHelper; 

use PHPExcel;
use PHPExcel_Reader_CSV;
use PHPExcel_Reader_Excel2007;
use PHPExcel_Reader_Excel5;
use PHPExcel_RichText;

class ExcelHelper
{
    /**
     * @var int
     */
    private static $error = 1;
    /**
     * @var string
     */
    private static $errorMsg = 'Nothing !';


    /**
     * @return array
     */
    public static function getError()
    {
        return [self::$error, self::$errorMsg];
    }


    /**
     * @param int $error
     * @param string $errorMsg
     */
    private static function setError(int $error, string $errorMsg)
    {
        self::$error    = $error;
        self::$errorMsg = $errorMsg;
    }

    /**
     * @param PHPExcel $objPHPExcel
     * @param string $pCoordinate
     * @return string
     * @throws \PHPExcel_Exception
     */
    private static function getTrueVal(PHPExcel $objPHPExcel, string $pCoordinate)
    {
        $val = $objPHPExcel->getActiveSheet()->getCell($pCoordinate)->getValue();
        if ($val instanceof PHPExcel_RichText) { //富文本转换字符串
            $val = $val->__toString();
        }
        return trim($val);
    }


    /**
     * @param $filePath
     * @param array $modelMap
     * @param array $uniqueKeys
     * @param \Closure $insertCallBack
     * @param string $replaceCallBack
     * @param array $replacementMap
     * @param string $colStartAt
     * @param int $rowStartAt
     * @throws \PHPExcel_Exception
     * @throws \PHPExcel_Reader_Exception
     */
    public static function impExcel2DB($filePath, array $modelMap, array $uniqueKeys, \Closure $insertCallBack, $replaceCallBack = '', array $replacementMap = [], $colStartAt = 'A', $rowStartAt = 1)
    {
        set_time_limit(0);

        if (!is_file($filePath)) {
            self::setError(1, "没有文件");
        }
        //$objReader = \PHPExcel_IOFactory::createReader('Excel5');
        $Excel_name = $filePath;
        // $objPHPExcel = $objReader->load($Excel_name,$encode='utf-8');
        $extension = strtolower(pathinfo($Excel_name, PATHINFO_EXTENSION));
        if ($extension === 'xlsx') {
            $objReader   = new PHPExcel_Reader_Excel2007();
            $objPHPExcel = $objReader->load($Excel_name);
        } else if ($extension === 'xls') {
            $objReader   = new PHPExcel_Reader_Excel5();
            $objPHPExcel = $objReader->load($Excel_name);
        } else if ($extension === 'csv') {
            $PHPReader = new PHPExcel_Reader_CSV();
            //默认输入字符集
            $PHPReader->setInputEncoding('GBK');
            //默认的分隔符
            $PHPReader->setDelimiter(',');
            //载入文件
            $objPHPExcel = $PHPReader->load($Excel_name);
        }
        $sheet = $objPHPExcel->getSheet(0);
        //验证格式是否标准
        $test = self::getTrueVal($objPHPExcel, $colStartAt . $rowStartAt);
        if (!in_array($test, $modelMap)) {
            self::setError(1, "表格数据格式错误,请下载导入模板");
        }

        $highestRow       = $sheet->getHighestRow(); // 取得总行数
        $highestColumn    = $sheet->getHighestColumn(); // 取得总列数
        $highestColumnNum = ord($highestColumn);
        $mapNameFlip      = array_flip($modelMap);
        $map              = [];

        for ($l = ord($colStartAt); $l <= $highestColumnNum; $l++) {
            $val = self::getTrueVal($objPHPExcel, chr($l) . $rowStartAt);
            if (in_array($val, $modelMap)) {
                $map[$mapNameFlip[$val]] = chr($l);
            }
        }
        $data = [];
        for ($i = $rowStartAt + 1; $i <= $highestRow; $i++) {
            foreach ($uniqueKeys as $uKey => $uVals) {
                $test2 = self::getTrueVal($objPHPExcel, $map[$uKey] . $i);
                if (empty($test2)) {
                    continue 2;
                }
                if (is_callable($replaceCallBack) && in_array($test2, $uVals)) {
                    $tmp2 = [];
                    foreach ($map as $key => $val) {
                        $tmp2[$key] = self::getTrueVal($objPHPExcel, $val . $i);
                        if (array_key_exists($key, $replacementMap)) {
                            $tmp2[$key] = (int)array_search($tmp2[$key], $replacementMap[$key]);
                        }
                    }
                    $replaceCallBack($tmp2);
                    continue 2;
                }
            }
            foreach ($map as $key => $val) {
                $tmp       = [];
                $tmp[$key] = self::getTrueVal($objPHPExcel, $val . $i);
                if (array_key_exists($key, $replacementMap)) {
                    $tmp[$key] = (int)array_search($tmp[$key], $replacementMap[$key]);
                }
            }
            if (!empty($tmp)) {
                $data[] = $tmp;
            }
        }
        if (!empty($data)) {
            $insertCallBack($data);
        }
        self::setError(0, 'SUCC');
    }
}
