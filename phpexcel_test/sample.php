<?php 
require_once 'vendor/phpoffice/phpexcel/Classes/PHPExcel.php';
require_once 'vendor/phpoffice/phpexcel/Classes/PHPExcel/IOFactory.php';
include_once "db.php";
// 取得日期範圍
$start_date = $_GET['start_date'] ?? date('Y-m-d');
$end_date = $_GET['end_date'] ?? (new DateTime())->modify('+10 days')->format('Y-m-d');
$m_header = generateDateRange($start_date, $end_date);

// 初始化 PHPExcel
$objPHPExcel = new PHPExcel();
$sheet = $objPHPExcel->getActiveSheet();

$startColumn = 'C';
$startColumnIndex = \PHPExcel_Cell::columnIndexFromString($startColumn);
$endColumnIndex = $startColumnIndex + count($m_header);
$row = 1;

$objPHPExcel->getActiveSheet()->getStyle('1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

	//匯出主要資料表
	$objPHPExcel->setActiveSheetIndex(0)
				->setCellValue('A1', '序號')
				->setCellValue('B1', '團隊')
				->setCellValue('C1', '員工')
				;

	//設置行列高度
	$objPHPExcel->getActiveSheet()->getRowDimension(1)->setRowHeight(25);

	//設置邊框線及顏色			
	$objPHPExcel->getActiveSheet()->getStyle('A1')->getBorders()->getAllBorders()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
	$objPHPExcel->getActiveSheet()->getStyle('A1')->getBorders()->getAllBorders()->getColor()->setRGB('000000');
	$objPHPExcel->getActiveSheet()->getStyle('B1')->getBorders()->getAllBorders()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
	$objPHPExcel->getActiveSheet()->getStyle('B1')->getBorders()->getAllBorders()->getColor()->setRGB('000000');
	$objPHPExcel->getActiveSheet()->getStyle('C1')->getBorders()->getAllBorders()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
	$objPHPExcel->getActiveSheet()->getStyle('C1')->getBorders()->getAllBorders()->getColor()->setRGB('000000');
	//設置垂直及水平對齊
	$objPHPExcel->getActiveSheet()->getStyle('A1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
	$objPHPExcel->getActiveSheet()->getStyle('A1')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
	$objPHPExcel->getActiveSheet()->getStyle('B1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
	$objPHPExcel->getActiveSheet()->getStyle('B1')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
	$objPHPExcel->getActiveSheet()->getStyle('C1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
	$objPHPExcel->getActiveSheet()->getStyle('C1')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
	//設置寬度
	$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(5);
	$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(20);
	$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(20);
	//設置字型大小
	$objPHPExcel->getActiveSheet()->getStyle('A1')->getFont()->setSize(12)->setBold(true);
	$objPHPExcel->getActiveSheet()->getStyle('B1')->getFont()->setSize(12)->setBold(true);
	$objPHPExcel->getActiveSheet()->getStyle('C1')->getFont()->setSize(12)->setBold(true);
	//設置底色
	$objPHPExcel->getActiveSheet()->getStyle('A1:C1')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
	$objPHPExcel->getActiveSheet()->getStyle('A1:C1')->getFill()->getStartColor()->setRGB('7FDBFF');

// 填入資料與樣式
foreach ($m_header as $i => $header) {
    $column = \PHPExcel_Cell::stringFromColumnIndex($startColumnIndex + $i);
    $cell = $column . $row;

    // 填入資料
    $sheet->setCellValue($cell, $header[1]);

    // 套用樣式
    applyCellStyles($sheet, $cell, 12);
}

function generateDateRange($start_date, $end_date) {
    $date1 = new DateTime($start_date);
    $date2 = new DateTime($end_date);

    $day_interval = new DateInterval('P1D');
    $period = new DatePeriod($date1, $day_interval, $date2->modify('+1 day'));

    $m_header = [];
    foreach ($period as $date) {
        $m_header[] = [$date->format('Y-m-d'), $date->format('d')];
    }
    return $m_header;
}

function applyCellStyles($sheet, $cell, $columnWidth = null, $backgroundColor = '7FDBFF') {
    // 設置邊框
    $sheet->getStyle($cell)->getBorders()->getAllBorders()
        ->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN)
        ->getColor()->setRGB('000000');
    
    // 設置對齊方式
    $sheet->getStyle($cell)->getAlignment()
        ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER)
        ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
    
    // 設置字型
    $sheet->getStyle($cell)->getFont()->setSize(12)->setBold(true);
    
    // 設置背景顏色
    $sheet->getStyle($cell)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
    $sheet->getStyle($cell)->getFill()->getStartColor()->setRGB($backgroundColor);
    
    // 設置欄位寬度
    if ($columnWidth !== null) {
        $sheet->getColumnDimensionByColumn(\PHPExcel_Cell::columnIndexFromString($cell[0]))->setWidth($columnWidth);
    }
}

// 添加 "合計" 標題
$summaryColumn = \PHPExcel_Cell::stringFromColumnIndex($endColumnIndex);
$summaryCell = $summaryColumn . $row;
$sheet->setCellValue($summaryCell, '合計');
applyCellStyles($sheet, $summaryCell, 20);

// 儲存檔案
$file = 'demo333.xlsx';
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save($file);

echo "檔案已儲存：{$file}";
// ?>