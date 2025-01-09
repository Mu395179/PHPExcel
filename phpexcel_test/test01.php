<?php
require_once 'vendor/phpoffice/phpexcel/Classes/PHPExcel.php';
require_once 'vendor/phpoffice/phpexcel/Classes/PHPExcel/IOFactory.php';

$objPHPExcel = new PHPExcel();
$sheet = $objPHPExcel->getActiveSheet();

$start_date = "2024-09-01";
$end_date = "2024-09-30";

// 總共天數
$days = Total_days($start_date, $end_date);

// 列出所有日期
show_days($days, $sheet);

// 儲存檔案
$file = 'demo_' . date('123') . '.xlsx';
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save($file);

echo "檔案已儲存：{$file}";

// 計算日期範圍的天數
function Total_days($start_date, $end_date)
{
    $start_timestamp = strtotime($start_date);
    $end_timestamp = strtotime($end_date);
    $days_difference = ceil(($end_timestamp - $start_timestamp) / (60 * 60 * 24)) + 1; // 加 1 包含開始日
    return $days_difference;
}

// 在工作表中填入日期
function show_days($days, $sheet)
{
    if ($sheet === null) {
        throw new Exception("Invalid sheet object.");
    }

    for ($i = 0; $i < $days; $i++) { // 從 0 開始計算
        $columnLetter = \PHPExcel_Cell::stringFromColumnIndex($i);
        $sheet->setCellValue($columnLetter . '1', (string)($i + 1)); // 數字從 1 開始
    }
}
