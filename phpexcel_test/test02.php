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
        'name' => 'Gura',
    ],
    [
        'id' => '2',
        'name' => 'Irys',
    ],
    [
        'id' => '3',
        'name' => 'Ame',
    ],
    [
        'id' => '4',
        'name' => 'Gigi',
    ],
];

$member_teamwork = [
    [
        'id' => '1',
        'dispatch_day' => '10',
        'attendance_status' => '支援',
        'attendance_time' => '2',
        'construction_site' => 'EN',
        'transition' => 'Y',
        'transition_team' => 'Ina',
    ],
    [
        'id' => '2',
        'dispatch_day' => '5',
        'attendance_status' => '支援',
        'attendance_time' => '2',
        'construction_site' => 'EN',
        'transition' => 'Y',
        'transition_team' => 'Kobo',
    ],
    [
        'id' => '3',
        'dispatch_day' => '2',
        'attendance_status' => '支援',
        'attendance_time' => '2',
        'construction_site' => 'EN',
        'transition' => 'Y',
        'transition_team' => 'Zeta',
    ],
    [
        'id' => '3',
        'dispatch_day' => '10',
        'attendance_status' => '支援',
        'attendance_time' => '2',
        'construction_site' => 'EN',
        'transition' => 'Y',
        'transition_team' => 'Kobo',
    ],
    [
        'id' => '4',
        'dispatch_day' => '12',
        'attendance_status' => '支援',
        'attendance_time' => '2',
        'construction_site' => 'EN',
        'transition' => 'Y',
        'transition_team' => 'Kobo',
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
        'B3' => '工地',
    ];
header_first($first_rows, $sheet);


// 最後列表頭欄位
header_last($days, $sheet);

// 列出所有日期
Excel_show_days($start_date,$days, $sheet);

// 員工資訊
member_info($member, $sheet);

// 支援報表
support_report($start_date, $end_date, $member_teamwork, $sheet);

// 儲存檔案
$file = 'demo_' . date('222') . '.xlsx';
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
    $columnLetter = \PHPExcel_Cell::stringFromColumnIndex(2 + $days);
    $sheet->mergeCells('A1:' . $columnLetter . '2');
    $sheet->setCellValue('A1', '工地狀況查詢');

    // 設定儲存格高度為40
    $sheet->getRowDimension(1)->setRowHeight(20);
    $sheet->getRowDimension(2)->setRowHeight(20);

    // 設定儲存格文字對齊
    $sheet->getStyle('A1')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    $sheet->getStyle('A1')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

    // 設定邊框
    $styleArray = [
        'borders' => [
            'allborders' => [
                'style' => PHPExcel_Style_Border::BORDER_THIN, // 粗線
                'color' => ['argb' => 'FF000000'], // 黑色
            ],
        ],
    ];
    //設置字型大小
    $sheet->getStyle('A1')->getFont()->setSize(30)->setBold(true);
    $sheet->getStyle('A1:' . $columnLetter . '2')->applyFromArray($styleArray);
}
// 首列表頭欄位
function header_first($first_rows, $sheet)
{
    foreach ($first_rows as $key => $value) {

        if ($value == 'A3') {
            $sheet->getColumnDimension('A')->setWidth(5);
        } else {
            $sheet->getColumnDimension('B')->setWidth(20);
        }
        // 設定儲存格值
        $sheet->setCellValue($key, $value);

        // 設定文字置中
        $sheet->getStyle($key)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle($key)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

        // 設定粗體字
        $sheet->getStyle($key)->getFont()->setBold(true);

        // 設定底色
        $sheet->getStyle($key)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
        $sheet->getStyle($key)->getFill()->getStartColor()->setRGB('7FDBFF'); // 底色為 #7FDBFF

        // 設定邊框
        $styleArray = [
            'borders' => [
                'allborders' => [
                    'style' => PHPExcel_Style_Border::BORDER_THIN, // 細線
                    'color' => ['argb' => 'FF000000'], // 黑色
                ],
            ],
        ];
        $sheet->getStyle($key)->applyFromArray($styleArray);
    }
}

// 最後列表頭欄位"合計"
function header_last($days, $sheet)
{
    $columnLetter = \PHPExcel_Cell::stringFromColumnIndex(2 + $days);

    // 設定文字置中
    $sheet->getStyle($columnLetter . '3')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    $sheet->getStyle($columnLetter . '3')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

    // 設定底色
    $sheet->getStyle($columnLetter . '3')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
    $sheet->getStyle($columnLetter . '3')->getFill()->getStartColor()->setRGB('7FDBFF'); // 底色為 #7FDBFF

    // 設定粗體字
    $sheet->getStyle($columnLetter . '3')->getFont()->setBold(true);

    // 設定邊框（全部細線）
    $styleBorders = [
        'borders' => [
            'allborders' => [
                'style' => PHPExcel_Style_Border::BORDER_THIN, // 細線
                'color' => ['argb' => 'FF000000'], // 黑色
            ],
        ],
    ];
    $sheet->getStyle($columnLetter . '3')->applyFromArray($styleBorders);

    // 設定內容為 "合計"
    $sheet->setCellValue($columnLetter . '3', '合計');

    // 設定列寬度為 20
    $sheet->getColumnDimension($columnLetter)->setWidth(20);
}


// 在工作表中填入日期
function Excel_show_days($start_date, $days, $sheet)
{
    $date = new DateTime($start_date);

    if ($sheet === null) {
        throw new Exception("Invalid sheet object.");
    }

    for ($i = 0; $i < $days; $i++) {
        $current_date = (clone $date)->modify("+$i days")->format('d'); // 每次加一天
        $columnLetter = PHPExcel_Cell::stringFromColumnIndex(2 + $i); // 計算欄位字母

        // 設定日期
        $sheet->setCellValue($columnLetter . '3', $current_date);

        // 設定文字置中
        $sheet->getStyle($columnLetter . '3')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle($columnLetter . '3')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

        // 設定底色
        $sheet->getStyle($columnLetter . '3')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
        $sheet->getStyle($columnLetter . '3')->getFill()->getStartColor()->setRGB('7FDBFF'); // 設置底色

        // 設定邊框（細線）
        $styleBorders = [
            'borders' => [
                'allborders' => [
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => ['argb' => 'FF000000'], // 黑色
                ],
            ],
        ];
        $sheet->getStyle($columnLetter . '3')->applyFromArray($styleBorders);
    }
}



//員工資訊 
function member_info($members, $sheet, $startRow = 4)
{
    // 計算成員數量
    $memberCount = count($members);

    for ($i = 0; $i < $memberCount; $i++) {
        $currentRow = $startRow + ($i * 6); // 每個成員佔用 6 行

        // 合併儲存格 (A 列，用於顯示 ID)
        $sheet->mergeCells("A{$currentRow}:A" . ($currentRow + 5));
        $sheet->setCellValue("A{$currentRow}", $members[$i]['id']);

        // 設置團隊 (B 列)
        $sheet->mergeCells("B{$currentRow}:B" . ($currentRow + 5));
        $sheet->setCellValue("B{$currentRow}", $members[$i]['name']);

        // 水平與垂直居中
        $sheet->getStyle("A{$currentRow}:B" . ($currentRow + 5))
            ->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle("A{$currentRow}:B" . ($currentRow + 5))
            ->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

        // 設置字體大小為 12
        $sheet->getStyle("A{$currentRow}:B" . ($currentRow + 5))
            ->getFont()->setSize(12);

        // 設置四邊框為黑色細線
        $sheet->getStyle("A{$currentRow}:B" . ($currentRow + 5))
        ->applyFromArray([
            'alignment' => [
                'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                'vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER,
            ],
            'borders' => [
                'allborders' => [
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => ['argb' => 'FF000000'],
                ],
            ],
            'font' => ['bold' => true],
        ]);
    }

    // 合計列
    $totalRow = $startRow + ($memberCount * 6); // 計算合計所在的行
    $sheet->mergeCells("A{$totalRow}:B{$totalRow}"); // 合併 A~B 列
    $sheet->setCellValue("A{$totalRow}", "合計"); // 設置值為「合計」

    // 設置居中對齊
    $sheet->getStyle("A{$totalRow}:B{$totalRow}")
        ->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    $sheet->getStyle("A{$totalRow}:B{$totalRow}")
        ->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

    // 設置加粗字體
    $sheet->getStyle("A{$totalRow}:B{$totalRow}")->getFont()->setBold(true);

    // 設置字體大小為 12
    $sheet->getStyle("A{$totalRow}:B{$totalRow}")
        ->getFont()->setSize(12);

    // 設置四邊框為黑色細線
    $sheet->getStyle("A{$totalRow}:B{$totalRow}")->applyFromArray([
        'alignment' => [
            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
            'vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER,
        ],
        'borders' => [
            'allborders' => [
                'style' => PHPExcel_Style_Border::BORDER_THICK,
                'color' => ['argb' => 'FF000000'],
            ],
        ],
        'font' => ['bold' => true],
    ]);

    // 設置合計列底色為 FFDC00
    $sheet->getStyle("A{$totalRow}:B{$totalRow}")
        ->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
    $sheet->getStyle("A{$totalRow}:B{$totalRow}")
        ->getFill()->getStartColor()->setRGB('FFDC00');
}
// 支援團隊報表
function support_report($start_date, $end_date, $member_teamwork, $sheet)
{
    if ($sheet === null) {
        throw new Exception("Invalid sheet object.");
    }

    // 計算總天數
    $total_days = Total_days($start_date, $end_date);

    // 開始日期的時間戳
    $start_timestamp = strtotime($start_date);

    // 整理資料：按 ID 與 dispatch_day 分組
    $teamwork_map = [];
    foreach ($member_teamwork as $entry) {
        $teamwork_map[$entry['id']][$entry['dispatch_day']] = $entry;
    }

    // 開始列偏移（假設日期從第 4 列開始）
    $column_offset = 2;
    $initial_row = 4; // 每個 ID 起始行
    $block_height = 3; // 每個 ID 的區塊高度

    // 計算最後一欄的字母
    $last_column_index = $column_offset + $total_days;
    $last_column_letter = \PHPExcel_Cell::stringFromColumnIndex($last_column_index);

    // 初始化每日加總數組
    $daily_totals = array_fill(1, $total_days, 0);

    foreach ($teamwork_map as $id => $dispatches) {
        $current_row = $initial_row;
        $total_hours = 0;

        foreach ($dispatches as $dispatch_day => $entry) {
            $columnIndex = $dispatch_day - 1; // 對應到日期的列
            $columnLetter = \PHPExcel_Cell::stringFromColumnIndex($column_offset + $columnIndex);

            // 設置 Attendance Time
            if (isset($entry['attendance_time'])) {
                $sheet->setCellValue("{$columnLetter}{$current_row}", $entry['attendance_time']);
                $total_hours += $entry['attendance_time']; // 累加總時數
                $daily_totals[$dispatch_day] += $entry['attendance_time']; // 每日加總
            }

            // 設置 Attendance Status 與 Construction Site
            if (isset($entry['attendance_status']) && $entry['attendance_status'] === "支援" && isset($entry['construction_site'])) {
                $sheet->setCellValue("{$columnLetter}" . ($current_row + 1), $entry['attendance_status'] . $entry['construction_site']);
            }

            // 設置 Transition 與 Transition Team
            if (isset($entry['transition']) && $entry['transition'] === "Y" && isset($entry['transition_team'])) {
                $sheet->setCellValue("{$columnLetter}" . ($current_row + 2), $entry['transition_team']);
            }
        }

        // 在最後一欄輸入總時數並合併儲存格（垂直兩格）
        $total_hours_cell = "{$last_column_letter}{$current_row}";
        $sheet->setCellValue($total_hours_cell, (string) ($total_hours));
        $sheet->mergeCells("{$total_hours_cell}:{$last_column_letter}" . ($current_row + 2)); // 合併下兩格

        // 每個 ID 的資料結束後，移動到下一區塊
        $initial_row += $block_height;
    }


    // 設置最後一列的每日加總
    $total_row = $initial_row; // 設定最後一列的位置
    for ($day = 1; $day <= $total_days; $day++) {
        $columnLetter = \PHPExcel_Cell::stringFromColumnIndex($column_offset + ($day - 1));
        $sheet->setCellValue("{$columnLetter}{$total_row}", (string) ($daily_totals[$day]));
        // 設置合計列底色為 FFDC00
        $sheet->getStyle("{$columnLetter}{$total_row}")
            ->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
        $sheet->getStyle("{$columnLetter}{$total_row}")
            ->getFill()->getStartColor()->setRGB('FFDC00');
    }

    // 設置儲存格樣式（判斷有無值）
    for ($row = 4; $row <= $total_row; $row++) {
        for ($col = $column_offset; $col <= $last_column_index; $col++) {
            $columnLetter = \PHPExcel_Cell::stringFromColumnIndex($col);
            $cellValue = $sheet->getCell("{$columnLetter}{$row}")->getValue();

            if ($cellValue) {
                // 儲存格有值
                $sheet->getColumnDimension($columnLetter)->setWidth(9);
            } else {
                // 儲存格無值
                $sheet->getColumnDimension($columnLetter)->setWidth(3);
            }

            $sheet->getRowDimension($row)->setRowHeight(20);

            // 設置樣式

            if($row===$total_row){
                $sheet->getStyle("{$columnLetter}{$row}")->applyFromArray([
                    'alignment' => [
                        'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                        'vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER,
                    ],
                    'borders' => [
                        'allborders' => [
                            'style' => PHPExcel_Style_Border::BORDER_THICK,
                            'color' => ['argb' => 'FF000000'],
                        ],
                    ],
                    'font' => ['bold' => true],
                ]);
            }else{
                $sheet->getStyle("{$columnLetter}{$row}")->applyFromArray([
                    'alignment' => [
                        'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
                        'vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER,
                    ],
                    'borders' => [
                        'allborders' => [
                            'style' => PHPExcel_Style_Border::BORDER_THIN,
                            'color' => ['argb' => 'FF000000'],
                        ],
                    ],
                    'font' => ['bold' => true],
                ]);
            }
            
        }
    }
}



function dd($array)
{
    echo "<pre>";
    print_r($array);
    echo "</pre>";
}
