<?php

namespace PHPExcelHelper;

use PHPExcel;
use PHPExcel_Reader_CSV;
use PHPExcel_Reader_Excel2007;
use PHPExcel_Reader_Excel5;
use PHPExcel_RichText;
use PHPExcel_Writer_Excel2007;
use PHPExcel_Writer_Excel5;

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
     * @param $filePath 文件路径
     * @param array $modelMap 字段名=>表头名
     * @param array $uniqueKeys 字段名 => [值...]
     * @param \Closure $insertCallBack 插入的数据 方法传入 从表格提取到的数据数组
     * @param string $replaceCallBack
     * @param array $replacementMap 唯一如果又重复 字段名=> [[替换值=>查找值]...] //TODO 完善重复后的替换
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

        for ($ord = ord($colStartAt); $ord <= $highestColumnNum; $ord++) {
            $val = self::getTrueVal($objPHPExcel, chr($ord) . $rowStartAt);
            if (in_array($val, $modelMap)) {
                $map[$mapNameFlip[$val]] = chr($ord);
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
                    foreach ($map as $field => $chr) {
                        $tmp2[$field] = self::getTrueVal($objPHPExcel, $chr . $i);
                        if (array_key_exists($field, $replacementMap)) {
                            $tmp2[$field] = (int)array_search($tmp2[$field], $replacementMap[$field]);
                        }
                    }
                    $replaceCallBack($tmp2);
                    continue 2;
                }
            }
            foreach ($map as $field => $chr) {
                $tmp       = [];
                $tmp[$field] = self::getTrueVal($objPHPExcel, $chr . $i);
                if (array_key_exists($field, $replacementMap)) {
                    $tmp[$field] = (int)array_search($tmp[$field], $replacementMap[$field]);
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

    protected static function excelClomunKeys()
    {
        $init =  [];
        $more_init = [];
        for($i = ord('A');$i<=ord('Z');$i++){
            $init[] = chr($i);
        }
        foreach ($init as $chr) {
            foreach ($init as $chr2) {
                $more_init[] = $chr . $chr2;
            }
        }
        return array_merge($init,$more_init);
    }


    /** 
     * 创建(导出)Excel数据表格 
     * @param  array   $list        要导出的数组格式的数据 
     * @param  string  $filename    导出的Excel表格数据表的文件名 
     * @param  array   $indexKey    $list数组中与Excel表格表头$header中每个项目对应的字段的名字(key值) 
     * @param  array   $startRow    第一条数据在Excel表格中起始行 
     * @param  [bool]  $excel2007   是否生成Excel2007(.xlsx)以上兼容的数据表 
     * 比如: $indexKey与$list数组对应关系如下: 
     *     $indexKey = array('id','username','sex','age'); 
     *     $list = array(array('id'=>1,'username'=>'YQJ','sex'=>'男','age'=>24)); 
     */
    static function exportExcel($list, $filename, $indexKey, $startRow = 1, $excel2007 = true)
    {
        //文件引入  

        if (empty($filename)) $filename = time();
        if (!is_array($indexKey)) return false;

        $header_arr = self::excelClomunKeys();

        //初始化PHPExcel()  
        $objPHPExcel = new PHPExcel();

        //设置保存版本格式  
        if ($excel2007) {
            $objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel);
            $filename = $filename . '.xlsx';
            $mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
        } else {
            $objWriter = new PHPExcel_Writer_Excel5($objPHPExcel);
            $filename = $filename . '.xls';
            $mime = 'application/vnd.ms-execl';
        }

        //接下来就是写数据到表格里面去  
        $objActSheet = $objPHPExcel->getActiveSheet();
        //$startRow = 1;  
        foreach ($list as $row) {
            foreach ($indexKey as $key => $value) {
                //这里是设置单元格的内容  
                $objActSheet->setCellValue($header_arr[$key] . $startRow, $row[$value]);
            }
            $startRow++;
        }

        // 下载这个表格，在浏览器输出  
        header("Pragma: public");
        header("Expires: 0");
        header("Cache-Control:must-revalidate, post-check=0, pre-check=0");
        header("Content-Type:application/force-download");
        header("Content-Type:".$mime);
        header("Content-Type:application/octet-stream");
        header("Content-Type:application/download");;
        header('Content-Disposition:attachment;filename=' . $filename . '');
        header("Content-Transfer-Encoding:binary");
        $objWriter->save('php://output');
    }
}
