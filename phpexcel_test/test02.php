<?php
require_once 'vendor/phpoffice/phpexcel/Classes/PHPExcel.php';
require_once 'vendor/phpoffice/phpexcel/Classes/PHPExcel/IOFactory.php';
// include_once 'db.php';

// 初始化Excel 物件
$objPHPExcel = new PHPExcel();
$sheet = $objPHPExcel->getActiveSheet();
// 時間範圍
$start_date = "2024-09-01";
$end_date = "2024-09-30";


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
function q($sql)
{
    $pdo = new PDO("mysql:host=localhost;charset=utf8;dbname=team_support", 'root', '');
    return $pdo->query($sql); // 返回 PDOStatement
}

// 執行查詢
$result = q('SELECT * FROM `dispatch`');

$members = []; // 初始化主陣列
$member_teamwork = []; // 用於存儲 work_day 資料

// 提取 dispatch 資料
while ($row = $result->fetch(2)) {
    $members[] = [
        'id' => $row['id'],
        'team' => $row['team_name'],
        'name' => $row['employee_name'],
        'employee_id' => $row['employee_id'],
    ];

    // 如果 employee_id 不為空，查詢對應的 work_day 資料
    if (!empty($row['employee_id'])) {
        $work_day_result = q("SELECT * FROM `work_day` WHERE `employee_id`='{$row['employee_id']}'");
        while ($row2 = $work_day_result->fetch(2)) {
            $member_teamwork[] = [
                'id' => $row2['id'],
                'dispatch_id' => $row2['dispatch_id'],
                'dispatch_year' => $row2['dispatch_year'],
                'dispatch_month' => $row2['dispatch_month'],
                'dispatch_day' => $row2['dispatch_day'], // 確保為整數
                'employee_id' => $row2['employee_id'],
                'manpower' => $row2['manpower'],
                'attendance_time' => $row2['attendance_day'], // 確保為浮點數
                'attendance_status' => $row2['attendance_status'], // 字串不需要轉換
                'transition' => $row2['transition'],
                'transition_start' => $row2['transition_start'],
                'transition_end' => $row2['transition_end'],
                'team_name' => $row2['team_name'],
                'transition_team_name' => $row2['transition_team_name'],
            ];
        }
    }
}



// 測試結果
dd(['members' => $members, 'member_teamwork' => $member_teamwork]);


// dd($members);
// 員工資訊
member_info($members, $sheet);

// 支援報表
support_report($start_date, $end_date, $member_teamwork, $sheet);

// 儲存檔案
$file = 'demo_' . date('666') . '.xlsx';
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
            $sheet->getColumnDimension('C')->setWidth(20);
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
    $columnLetter = \PHPExcel_Cell::stringFromColumnIndex(3 + $days);

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
function Excel_show_days($days, $sheet)
{
    if ($sheet === null) {
        throw new Exception("Invalid sheet object.");
    }

    for ($i = 0; $i < $days; $i++) { // 從 0 開始計算
        $columnLetter = \PHPExcel_Cell::stringFromColumnIndex(3 + $i);

        // 設定日期
        $sheet->setCellValue($columnLetter . '3', (string) ($i + 1)); // 數字從 1 開始

        // 設定文字置中
        $sheet->getStyle($columnLetter . '3')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle($columnLetter . '3')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

        // 設定底色
        $sheet->getStyle($columnLetter . '3')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
        $sheet->getStyle($columnLetter . '3')->getFill()->getStartColor()->setRGB('7FDBFF'); // 底色為 #7FDBFF

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
    }
}



//員工資訊 
function member_info($members, $sheet, $startRow = 4)
{
    // 計算成員數量
    $memberCount = count($members);

    for ($i = 0; $i < $memberCount; $i++) {
        $currentRow = $startRow + ($i * 3); // 每個成員佔用 3 行

        // 合併儲存格 (A 列，用於顯示 ID)
        $sheet->mergeCells("A{$currentRow}:A" . ($currentRow + 2));
        $sheet->setCellValue("A{$currentRow}", $members[$i]['id']);

        // 設置團隊 (B 列)
        $sheet->mergeCells("B{$currentRow}:B" . ($currentRow + 2));
        $sheet->setCellValue("B{$currentRow}", $members[$i]['team']);

        // 設置成員名稱 (C 列)
        $sheet->mergeCells("C{$currentRow}:C" . ($currentRow + 2));
        $sheet->setCellValue("C{$currentRow}", $members[$i]['name']);

        // 水平與垂直居中
        $sheet->getStyle("A{$currentRow}:C" . ($currentRow + 2))
            ->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle("A{$currentRow}:C" . ($currentRow + 2))
            ->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

        // 設置字體大小為 12
        $sheet->getStyle("A{$currentRow}:C" . ($currentRow + 2))
            ->getFont()->setSize(12);

        // 設置四邊框為黑色細線
        $sheet->getStyle("A{$currentRow}:C" . ($currentRow + 2))
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
    $totalRow = $startRow + ($memberCount * 3); // 計算合計所在的行
    $sheet->mergeCells("A{$totalRow}:C{$totalRow}"); // 合併 A~C 列
    $sheet->setCellValue("A{$totalRow}", "合計"); // 設置值為「合計」

    // 設置居中對齊
    $sheet->getStyle("A{$totalRow}:C{$totalRow}")
        ->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    $sheet->getStyle("A{$totalRow}:C{$totalRow}")
        ->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

    // 設置加粗字體
    $sheet->getStyle("A{$totalRow}:C{$totalRow}")->getFont()->setBold(true);

    // 設置字體大小為 12
    $sheet->getStyle("A{$totalRow}:C{$totalRow}")
        ->getFont()->setSize(12);

    // 設置四邊框為黑色細線
    $sheet->getStyle("A{$totalRow}:C{$totalRow}")->applyFromArray([
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
    $sheet->getStyle("A{$totalRow}:C{$totalRow}")
        ->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
    $sheet->getStyle("A{$totalRow}:C{$totalRow}")
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
    $column_offset = 3;
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
            if ($dispatch_day >= 1 && $dispatch_day <= $total_days) {
                $columnIndex = $dispatch_day - 1; // 對應到日期的列
                $columnLetter = \PHPExcel_Cell::stringFromColumnIndex($column_offset + $columnIndex);
    
                // 設置 Attendance Time
                if (isset($entry['attendance_time'])) {
                    $sheet->setCellValue("{$columnLetter}{$current_row}", $entry['attendance_time']);
                    $total_hours += $entry['attendance_time']; // 累加總時數
                    $daily_totals[$columnIndex] += $entry['attendance_time']; // 每日加總
                }
    
                // 設置 Attendance Status 與 Construction Site
                if (isset($entry['attendance_status']) && $entry['attendance_status'] === "支援" && isset($entry['team_name'])) {
                    $sheet->setCellValue("{$columnLetter}" . ($current_row + 1), $entry['attendance_status'] . $entry['team_name']);
                }
    
                // 設置 Transition 與 Transition Team
                if (isset($entry['transition']) && $entry['transition'] === "Y" && isset($entry['transition_team'])) {
                    $sheet->setCellValue("{$columnLetter}" . ($current_row + 2), $entry['transition_team']);
                }
            } else {
                error_log("Invalid dispatch_day: {$dispatch_day}");
            }
        }
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

            if ($row === $total_row) {
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
            } else {
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
