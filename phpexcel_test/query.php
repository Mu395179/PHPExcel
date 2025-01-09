<?php
require_once 'vendor/phpoffice/phpexcel/Classes/PHPExcel.php';
require_once 'vendor/phpoffice/phpexcel/Classes/PHPExcel/IOFactory.php';
include_once "db.php";

$objPHPExcel = new PHPExcel();
$sheet = $objPHPExcel->getActiveSheet();
$startColumn = 'C';
$startColumnIndex = \PHPExcel_Cell::columnIndexFromString($startColumn);
$endColumnIndex = $startColumnIndex + 30;
$row = 1;

$objPHPExcel->getActiveSheet()->getStyle('1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

// 匯出主要資料表
$objPHPExcel->setActiveSheetIndex(0)
    ->setCellValue('A1', '序號')
    ->setCellValue('B1', '團隊')
    ->setCellValue('C1', '員工');

    $objPHPExcel->getActiveSheet()->getRowDimension(1)->setRowHeight(25);

    // 設置字體大小和粗體
    $objPHPExcel->getActiveSheet()->getStyle('A1:C1')->getFont()->setSize(12)->setBold(true);
    
    // 設置背景填充顏色
    $objPHPExcel->getActiveSheet()->getStyle('A1:C1')->getFill()
        ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)
        ->getStartColor()->setRGB('7FDBFF');
    
    // 增加黑色邊框
    $objPHPExcel->getActiveSheet()->getStyle('A1:C1')->getBorders()->getAllBorders()
        ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN)
        ->getColor()->setRGB('000000');
    
    // 設置文字水平和垂直置中
    $objPHPExcel->getActiveSheet()->getStyle('A1:C1')->getAlignment()
        ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER)
        ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

// 填入資料
for ($i = $startColumnIndex; $i <= $endColumnIndex; $i++) {
    $column = \PHPExcel_Cell::stringFromColumnIndex($i); // 獲取列名
    $cell = $column . ($row); // 填入第2行
    $sheet->setCellValue($cell, $i - $startColumnIndex + 1); // 從 1 開始填入數據
    applyCellStyles($sheet, $cell, 12);
}

// 添加 "合計" 標題
$summaryColumn = \PHPExcel_Cell::stringFromColumnIndex($endColumnIndex + 1);
$summaryCell = $summaryColumn . ($row);
$sheet->setCellValue($summaryCell, '合計');
applyCellStyles($sheet, $summaryCell, 20);

// 儲存檔案
$file = __DIR__ . '/demo444.xlsx';
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save($file);

echo "檔案已儲存：{$file}";

// 套用樣式函數
function applyCellStyles($sheet, $cell, $columnWidth = null, $backgroundColor = '7FDBFF') {
    $sheet->getStyle($cell)->getBorders()->getAllBorders()
        ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN)
        ->getColor()->setRGB('000000');
    $sheet->getStyle($cell)->getAlignment()
        ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER)
        ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
    $sheet->getStyle($cell)->getFont()->setSize(12)->setBold(true);
    $sheet->getStyle($cell)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
    $sheet->getStyle($cell)->getFill()->getStartColor()->setRGB($backgroundColor);

    if ($columnWidth !== null) {
        $sheet->getColumnDimension($cell[0])->setWidth($columnWidth);
    }
}
?>
