<?php
require_once 'vendor/phpoffice/phpexcel/Classes/PHPExcel.php';
require_once 'vendor/phpoffice/phpexcel/Classes/PHPExcel/IOFactory.php';
// include_once 'db.php';

// 初始化Excel 物件
$objPHPExcel = new PHPExcel();
$sheet = $objPHPExcel->getActiveSheet();

// 時間範圍
$start_date = "2024-09-01";
$end_date = "2024-09-15";

$member = [
    [
        'id' => '1',
        'team' => 'holoEn',
        'name' => 'Gura',
    ],
    [
        'id' => '2',
        'team' => 'holoEn',
        'name' => 'Irys',
    ],
    [
        'id' => '3',
        'team' => 'holoEn',
        'name' => 'Ame',
    ],
    [
        'id' => '4',
        'team' => 'holoEn',
        'name' => 'Gigi',
    ],
];


// 總共天數
$days = Total_days($start_date, $end_date);

// 標題
table_title($days, $sheet);

// 首列表頭欄位
$first_rows =
    [
        'A3' => '序號',
        'B3' => '團隊',
        'C3' => '員工'
    ];
header_first($first_rows, $sheet);


// 最後列表頭欄位
header_last($days, $sheet);

// 列出所有日期
Excel_show_days($days, $sheet);

// 儲存檔案
$file = 'demo_' . date('555') . '.xlsx';
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

// 標題
function table_title($days, $sheet)
{
    $columnLetter = \PHPExcel_Cell::stringFromColumnIndex(3 + $days);
    $sheet->mergeCells('A1:' . $columnLetter . '2');
    $sheet->setCellValue('A1', '團隊報表');
    $sheet->getStyle('A1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    $sheet->getStyle('A1')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
}

// 首列表頭欄位
function header_first($first_rows, $sheet)
{
    foreach ($first_rows as $key => $value) {
        $sheet->setCellValue($key, $value);
    }

}


// 在工作表中填入日期
function Excel_show_days($days, $sheet)
{
    if ($sheet === null) {
        throw new Exception("Invalid sheet object.");
    }

    for ($i = 0; $i < $days; $i++) { // 從 0 開始計算
        $columnLetter = \PHPExcel_Cell::stringFromColumnIndex(3 + $i);
        $sheet->setCellValue($columnLetter . '3', (string) ($i + 1)); // 數字從 1 開始
    }
}

// 最後列表頭欄位"合計"
function header_last($days, $sheet)
{
    $columnLetter = \PHPExcel_Cell::stringFromColumnIndex(3 + $days);
    $sheet->setCellValue($columnLetter . '3', '合計');
}

function dd($array)
{
    echo "<pre>";
    print_r($array);
    echo "</pre>";
}
