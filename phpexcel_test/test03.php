<?php
require_once 'vendor/phpoffice/phpexcel/Classes/PHPExcel.php';
require_once 'vendor/phpoffice/phpexcel/Classes/PHPExcel/IOFactory.php';
// include_once 'db.php';

// 初始化Excel 物件
$objPHPExcel = new PHPExcel();
$sheet = $objPHPExcel->getActiveSheet();




// 時間範圍
$start_date = "2024-09-01";
$end_date = "2024-09-10";

$member = [
    [
        'id' => '1',
        'company' => '台積電',
    ],
    [
        'id' => '2',
        'company' => '華碩',
    ],
    [
        'id' => '3',
        'company' => '台達電',
    ],
    [
        'id' => '4',
        'company' => '聯發科',
    ],
];

$member_teamwork = [

    [
        'id' => '1',
        'dispatch_company' => '華碩',
        'team_id' => '2024A',
        'team_name' => '電工C組',
        'dispatch_day' => '2',
        'construction_site' => '新竹廠',
        'manpower' => '1',
        'workinghours' => '1',
        'manpower_supported' => '1',
        'workinghours_supported' => '1',
        'foreign_manpower_supported' => '1',
        'foreign_workinghours_supported' => '1',
    ],

    [
        'id' => '2',
        'dispatch_company' => '華碩',
        'team_id' => '2024A',
        'team_name' => '電工A組',
        'dispatch_day' => '3',
        'construction_site' => '新竹廠',
        'manpower' => '2',
        'workinghours' => '2',
        'manpower_supported' => '2',
        'workinghours_supported' => '2',
        'foreign_manpower_supported' => '2',
        'foreign_workinghours_supported' => '2',
    ],
    [
        'id' => '2',
        'dispatch_company' => '華碩',
        'team_id' => '2024A',
        'team_name' => '電工A組',
        'dispatch_day' => '5',
        'construction_site' => '新竹廠',
        'manpower' => '6',
        'workinghours' => '6',
        'manpower_supported' => '6',
        'workinghours_supported' => '6',
        'foreign_manpower_supported' => '6',
        'foreign_workinghours_supported' => '6',
    ],
    [
        'id' => '2',
        'dispatch_company' => '華碩',
        'team_id' => '2024A',
        'team_name' => '電工A組',
        'dispatch_day' => '4',
        'construction_site' => '新竹廠',
        'manpower' => '3',
        'workinghours' => '3',
        'manpower_supported' => '3',
        'workinghours_supported' => '3',
        'foreign_manpower_supported' => '3',
        'foreign_workinghours_supported' => '3',
    ],
    [
        'id' => '3',
        'dispatch_company' => '台達電',
        'team_id' => '2024A',
        'team_name' => '電工A組',
        'dispatch_day' => '6',
        'construction_site' => '新竹廠',
        'manpower' => '4',
        'workinghours' => '4',
        'manpower_supported' => '4',
        'workinghours_supported' => '4',
        'foreign_manpower_supported' => '4',
        'foreign_workinghours_supported' => '4',
    ],
    [
        'id' => '4',
        'dispatch_company' => '聯發科',
        'team_id' => '2024A',
        'team_name' => '電工A組',
        'dispatch_day' => '8',
        'construction_site' => '新竹廠',
        'manpower' => '5',
        'workinghours' => '5',
        'manpower_supported' => '5',
        'workinghours_supported' => '5',
        'foreign_manpower_supported' => '5',
        'foreign_workinghours_supported' => '5',
    ],

];

// foreach ($member_teamwork as $key => $value) {

//         $workmake[$value['id']['dispatch_day']]=$value;
//         dd($workmake);
// }




// 總共天數
$days = Total_days($start_date, $end_date);

// 標題
table_title($days, $sheet);

// 最後列表頭欄位
header_last($days, $sheet);

// 列出所有日期
Excel_show_days($start_date, $days, $sheet);

// 首列表頭欄位
$first_rows =
    [
        'A3' => '序號',
        'B3' => '工地',
    ];
header_first($first_rows, $sheet);

// 員工資訊
member_info($member, $sheet);

// 支援報表
support_report($start_date, $end_date, $member_teamwork, $sheet);

// 儲存檔案
$file = 'demo_' . date('333') . '.xlsx';
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

        // 計算起始和結束欄位字母
        $startColumnLetter = PHPExcel_Cell::stringFromColumnIndex(2 + $i * 4);
        $endColumnLetter = PHPExcel_Cell::stringFromColumnIndex(2 + $i * 4 + 3);

        // 合併儲存格
        $mergedCells = $startColumnLetter . '3:' . $endColumnLetter . '3';
        $sheet->mergeCells($mergedCells);

        // 設定日期
        $sheet->setCellValue($startColumnLetter . '3', $current_date);

        // 設定文字置中
        $sheet->getStyle($mergedCells)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle($mergedCells)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

        // 設定底色
        $sheet->getStyle($mergedCells)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
        $sheet->getStyle($mergedCells)->getFill()->getStartColor()->setRGB('7FDBFF'); // 設置底色

        // 設定邊框（細線）
        $styleBorders = [
            'borders' => [
                'allborders' => [
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => ['argb' => 'FF000000'], // 黑色
                ],
            ],
        ];
        $sheet->getStyle($mergedCells)->applyFromArray($styleBorders);
    }
    // 新增最後的「合計」欄位
    $finalColumnLetter = PHPExcel_Cell::stringFromColumnIndex(2 + $days * 4);

    // 設定「合計」文字
    $sheet->setCellValue($finalColumnLetter . '3', '合計');

    // 設定文字置中
    $sheet->getStyle($finalColumnLetter . '3')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    $sheet->getStyle($finalColumnLetter . '3')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

    // 設定底色
    $sheet->getStyle($finalColumnLetter . '3')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
    $sheet->getStyle($finalColumnLetter . '3')->getFill()->getStartColor()->setRGB('FFD700'); // 金色

    // 設定邊框（細線）
    $sheet->getStyle($finalColumnLetter . '3')->applyFromArray([
        'borders' => [
            'allborders' => [
                'style' => PHPExcel_Style_Border::BORDER_THIN,
                'color' => ['argb' => 'FF000000'], // 黑色
            ],
        ],
    ]);

    // 設定所有欄位寬度
    for ($col = 0; $col <= (2 + $days * 4 + 1); $col++) {
        $sheet->getColumnDimensionByColumn($col)->setWidth(6);
    }
}

function form_style($sheet, $start_date,$days,$member_count){
    $date = new DateTime($start_date);

    if ($sheet === null) {
        throw new Exception("Invalid sheet object.");
    }

    for ($i = 0; $i < $days; $i++) {
        $current_date = (clone $date)->modify("+$i days")->format('d'); // 每次加一天

        // 計算起始和結束欄位字母
        $startColumnLetter = PHPExcel_Cell::stringFromColumnIndex(2 + $i * 4);
        $endColumnLetter = PHPExcel_Cell::stringFromColumnIndex(2 + $i * 4 + 3);

        // 合併儲存格
        $mergedCells = $startColumnLetter . '3:' . $endColumnLetter . '3';
        $sheet->mergeCells($mergedCells);

        // 設定日期
        $sheet->setCellValue($startColumnLetter . '3', $current_date);

        // 設定文字置中
        $sheet->getStyle($mergedCells)->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle($mergedCells)->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

        // 設定底色
        $sheet->getStyle($mergedCells)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
        $sheet->getStyle($mergedCells)->getFill()->getStartColor()->setRGB('7FDBFF'); // 設置底色

        // 設定邊框（細線）
        $styleBorders = [
            'borders' => [
                'allborders' => [
                    'style' => PHPExcel_Style_Border::BORDER_THIN,
                    'color' => ['argb' => 'FF000000'], // 黑色
                ],
            ],
        ];
        $sheet->getStyle($mergedCells)->applyFromArray($styleBorders);
    }
    // 新增最後的「合計」欄位
    $finalColumnLetter = PHPExcel_Cell::stringFromColumnIndex(2 + $days * 4);

    // 設定「合計」文字
    $sheet->setCellValue($finalColumnLetter . '3', '合計');

    // 設定文字置中
    $sheet->getStyle($finalColumnLetter . '3')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    $sheet->getStyle($finalColumnLetter . '3')->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

    // 設定底色
    $sheet->getStyle($finalColumnLetter . '3')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
    $sheet->getStyle($finalColumnLetter . '3')->getFill()->getStartColor()->setRGB('FFD700'); // 金色

    // 設定邊框（細線）
    $sheet->getStyle($finalColumnLetter . '3')->applyFromArray([
        'borders' => [
            'allborders' => [
                'style' => PHPExcel_Style_Border::BORDER_THIN,
                'color' => ['argb' => 'FF000000'], // 黑色
            ],
        ],
    ]);
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
        $sheet->setCellValue("B{$currentRow}", $members[$i]['company']);

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

    // A欄垂直合併兩格
    $sheet->mergeCells("A{$totalRow}:A" . ($totalRow + 1));
    $sheet->setCellValue("A{$totalRow}", "合計"); // 設置值為「合計」

    // B欄第一格設置總人數
    $sheet->setCellValue("B{$totalRow}", "總人數");

    // B欄第二格設置總工數
    $totalWork = $memberCount * 6; // 假設每個成員佔用 6 行即為總工數
    $sheet->setCellValue("B" . ($totalRow + 1), "總工數");

    // 設置底色為黑色 (FFDC00)
    $sheet->getStyle("A{$totalRow}:B" . ($totalRow + 1))
        ->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
    $sheet->getStyle("A{$totalRow}:B" . ($totalRow + 1))
        ->getFill()->getStartColor()->setRGB('FFDC00');

    // 設置居中對齊
    $sheet->getStyle("A{$totalRow}:B" . ($totalRow + 1))
        ->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    $sheet->getStyle("A{$totalRow}:B" . ($totalRow + 1))
        ->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);

    // 設置加粗字體
    $sheet->getStyle("A{$totalRow}:B" . ($totalRow + 1))->getFont()->setBold(true);

    // 設置字體大小為 12
    $sheet->getStyle("A{$totalRow}:B" . ($totalRow + 1))
        ->getFont()->setSize(12);

    // 設置四邊框為黑色細線
    $sheet->getStyle("A{$totalRow}:B" . ($totalRow + 1))->applyFromArray([
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
}
// 支援團隊報表
function support_report($start_date, $end_date, $member_teamwork, $sheet)
{
    if ($sheet === null) {
        throw new Exception("Invalid sheet object.");
    }

    // 計算總天數
    $total_days = Total_days($start_date, $end_date);

    // 整理資料：按 ID 與 dispatch_day 分組
    $teamwork_map = [];
    foreach ($member_teamwork as $entry) {
        $teamwork_map[$entry['id']][$entry['dispatch_day']] = $entry;
    }

    // 開始列偏移（假設日期從第 4 列開始）
    $column_offset = 2; // 起始於 C 欄
    $initial_row = 4; // 每組數據的初始行
    $block_height = 6; // 每組數據佔用行數

    // 計算最後一欄的字母
    $last_column_index = $column_offset + $total_days * 4;
    $last_column_letter = \PHPExcel_Cell::stringFromColumnIndex($last_column_index);

    // 初始化每日加總數組
    $daily_totals_manpower = array_fill(1, $total_days, 0);
    $daily_totals_workinghours = array_fill(1, $total_days, 0);
    foreach ($teamwork_map as $id => $dispatches) {
        $total_manpower = 0;
        $total_hours = 0;

        foreach ($dispatches as $dispatch_day => $entry) {
            $entry = $dispatches[$dispatch_day] ?? null; // 取出當天的數據，沒有則為 null
            $current_row = $initial_row;
            $merge_end_row = $current_row + 3;
            $column_base = ($dispatch_day - 1) * 4; // 基礎列索引
            // 計算各列字母
            $column_team_start = \PHPExcel_Cell::stringFromColumnIndex($column_offset + $column_base);
            $column_team_mid_1 = \PHPExcel_Cell::stringFromColumnIndex($column_offset + $column_base + 1);
            $column_team_mid_2 = \PHPExcel_Cell::stringFromColumnIndex($column_offset + $column_base + 2);
            $column_team_end = \PHPExcel_Cell::stringFromColumnIndex($column_offset + $column_base + 3);

            // 定義邊框樣式
            $styleArray = [
                'borders' => [
                    'allborders' => [
                        'style' => PHPExcel_Style_Border::BORDER_THIN, // 細線
                        'color' => ['rgb' => '000000'], // 黑色
                    ],
                ],
                'alignment' => [
                    'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER, // 文字水平置中
                    'vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER, // 文字垂直置中
                ],
            ];
            $start_row = $current_row; // 記錄起始行

                if (isset($entry['team_name'])) {
                    // 設置 team_name
                    $sheet->mergeCells("{$column_team_start}{$current_row}:{$column_team_end}{$current_row}");
                    $sheet->setCellValue("{$column_team_start}{$current_row}", $entry['team_name'] . "小隊");
                    // 套用樣式到合併的儲存格
                    $sheet->getStyle("{$column_team_start}{$current_row}:{$column_team_end}{$current_row}")->applyFromArray($styleArray);
                    $current_row++;
                    // 設置 manpower 與 workinghours
                    if (isset($entry['manpower']) && isset($entry['workinghours'])) {
                        $sheet->setCellValue("{$column_team_start}{$current_row}", "人數");
                        $sheet->setCellValue("{$column_team_mid_1}{$current_row}", $entry['manpower']);
                        $sheet->setCellValue("{$column_team_mid_2}{$current_row}", "工數");
                        $sheet->setCellValue("{$column_team_end}{$current_row}", $entry['workinghours']);
                        // 套用邊框樣式
                        $sheet->getStyle("{$column_team_start}{$current_row}:{$column_team_end}{$current_row}")->applyFromArray($styleArray);
                        $total_manpower += $entry['manpower'];
                        $total_hours += $entry['workinghours'];
                        $daily_totals_manpower[$dispatch_day] += $entry['manpower']; // 每日加總人工數
                        $daily_totals_workinghours[$dispatch_day] += $entry['workinghours']; // 每日加總工時數

                        $current_row++;

                    }

                    // 設置 foreign_manpower_supported 與 foreign_workinghours_supported
                    if (isset($entry['foreign_manpower_supported']) && isset($entry['foreign_workinghours_supported'])) {
                        $sheet->mergeCells("{$column_team_start}{$current_row}:{$column_team_end}{$current_row}");
                        $sheet->getStyle("{$column_team_start}{$current_row}:{$column_team_end}{$current_row}")->applyFromArray($styleArray);
                        $sheet->setCellValue("{$column_team_start}{$current_row}", "被支援移工");
                        // 套用邊框樣式
                        $sheet->getStyle("{$column_team_start}{$current_row}:{$column_team_end}{$current_row}")->applyFromArray($styleArray);
                        $current_row++;

                        $sheet->setCellValue("{$column_team_start}{$current_row}", "人數");
                        $sheet->setCellValue("{$column_team_mid_1}{$current_row}", $entry['foreign_manpower_supported']);
                        $sheet->setCellValue("{$column_team_mid_2}{$current_row}", "工數");
                        $sheet->setCellValue("{$column_team_end}{$current_row}", $entry['foreign_workinghours_supported']);
                        // 套用邊框樣式
                        $sheet->getStyle("{$column_team_start}{$current_row}:{$column_team_end}{$current_row}")->applyFromArray($styleArray);
                        $total_manpower += $entry['foreign_manpower_supported'];
                        $total_hours += $entry['foreign_workinghours_supported'];
                        $daily_totals_manpower[$dispatch_day] += $entry['foreign_manpower_supported']; // 每日加總人工數
                        $daily_totals_workinghours[$dispatch_day] += $entry['foreign_workinghours_supported']; // 每日加總工時數
                        $current_row++;
                    }

                    // 設置 foreign_manpower_supported 與 foreign_workinghours_supported
                    if (isset($entry['manpower_supported']) && isset($entry['workinghours_supported'])) {
                        $sheet->mergeCells("{$column_team_start}{$current_row}:{$column_team_end}{$current_row}");
                        $sheet->setCellValue("{$column_team_start}{$current_row}", "被支援外調");
                        // 套用邊框樣式
                        $sheet->getStyle("{$column_team_start}{$current_row}:{$column_team_end}{$current_row}")->applyFromArray($styleArray);
                        $current_row++;

                        $sheet->setCellValue("{$column_team_start}{$current_row}", "人數");
                        $sheet->setCellValue("{$column_team_mid_1}{$current_row}", $entry['manpower_supported']);
                        $sheet->setCellValue("{$column_team_mid_2}{$current_row}", "工數");
                        $sheet->setCellValue("{$column_team_end}{$current_row}", $entry['workinghours_supported']);
                        // 套用邊框樣式
                        $sheet->getStyle("{$column_team_start}{$current_row}:{$column_team_end}{$current_row}")->applyFromArray($styleArray);
                        $total_manpower += $entry['manpower_supported'];
                        $total_hours += $entry['workinghours_supported'];
                        $daily_totals_manpower[$dispatch_day] += $entry['manpower_supported']; // 每日加總人工數
                        $daily_totals_workinghours[$dispatch_day] += $entry['workinghours_supported']; // 每日加總工時數
                        $current_row++;
                    }
                }else{
                    dd($entry);
                }

        }
        // 在最後一欄輸入總時數
        $last_row_index = $current_row - 6;
        $total_hours_cell = "{$last_column_letter}{$last_row_index}";
        $sheet->setCellValue($total_hours_cell, "總時數:" . (string) $total_hours . "/n總工時:" . $total_manpower);

        // 計算合併範圍（從 $lest_row_index 向下 6 列）
        $merge_start = "{$last_column_letter}{$last_row_index}";
        $merge_end = "{$last_column_letter}" . ($last_row_index + 5);
        $merge_range = "{$merge_start}:{$merge_end}";
        // 設定垂直置中
        $sheet->getStyle($merge_range)->getAlignment()->setVertical(\PHPExcel_Style_Alignment::VERTICAL_CENTER);

        // 設定黑色細線邊框
        $sheet->getStyle($merge_range)->applyFromArray($styleArray);

        // 合併儲存格
        $sheet->mergeCells($merge_range);

        // 設定最後一欄的寬度增加 10
        $columnWidth = $sheet->getColumnDimension($last_column_letter)->getWidth();
        $sheet->getColumnDimension($last_column_letter)->setWidth($columnWidth + 1);

        // 如果 $total_hours 和 $total_manpower 皆為 0 或空值，不填入內容
        if (!empty($total_hours) || !empty($total_manpower)) {
            $total_hours_cell = "{$last_column_letter}{$last_row_index}";
            $sheet->setCellValue($total_hours_cell, "總時數:" . (string) $total_hours . "\n 總工數:" . (string) $total_manpower);
            $sheet->getStyle($total_hours_cell)->getAlignment()->setWrapText(true);
        }

        // 每個 ID 的資料結束後，移動到下一區塊
        $initial_row += $block_height;
    }



    // 設置最後一列的每日加總
    $total_row = $initial_row; // 設定最後一列的位置
    $merge_width = 4; // 每格合併 4 個儲存格

    $borderStyle = array(
        'borders' => array(
            'allborders' => array(
                'style' => PHPExcel_Style_Border::BORDER_MEDIUM, // 粗邊框
                'color' => array('rgb' => '000000') // 黑色
            )
        )
    );

    $alignmentStyle = array(
        'alignment' => array(
            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER, // 水平置中
            'vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER // 垂直置中
        )
    );

    $fontStyle = array(
        'font' => array(
            'bold' => true // 設為粗體
        )
    );

    for ($day = 1; $day <= $total_days; $day++) {
        $startColumnIndex = $column_offset + ($day - 1) * $merge_width;
        $startColumnLetter = PHPExcel_Cell::stringFromColumnIndex($startColumnIndex);
        $endColumnLetter = PHPExcel_Cell::stringFromColumnIndex($startColumnIndex + $merge_width - 1);

        // **合併每日總人數儲存格**
        $mergedCellManpower = "{$startColumnLetter}{$total_row}:{$endColumnLetter}{$total_row}";
        $sheet->mergeCells($mergedCellManpower);
        $sheet->setCellValue("{$startColumnLetter}{$total_row}", (string) ($daily_totals_manpower[$day]));

        // **合併每日總工時儲存格**
        $mergedCellHours = "{$startColumnLetter}" . ($total_row + 1) . ":{$endColumnLetter}" . ($total_row + 1);
        $sheet->mergeCells($mergedCellHours);
        $sheet->setCellValue("{$startColumnLetter}" . ($total_row + 1), (string) ($daily_totals_workinghours[$day]));

        // **設定背景顏色 (FFDC00)**
        $sheet->getStyle($mergedCellManpower)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
        $sheet->getStyle($mergedCellManpower)->getFill()->getStartColor()->setRGB('FFDC00');

        $sheet->getStyle($mergedCellHours)->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
        $sheet->getStyle($mergedCellHours)->getFill()->getStartColor()->setRGB('FFDC00');

        // **設定邊框**
        $sheet->getStyle($mergedCellManpower)->applyFromArray($borderStyle);
        $sheet->getStyle($mergedCellHours)->applyFromArray($borderStyle);

        // **設定文字置中**
        $sheet->getStyle($mergedCellManpower)->applyFromArray($alignmentStyle);
        $sheet->getStyle($mergedCellHours)->applyFromArray($alignmentStyle);

        // **設定文字粗體**
        $sheet->getStyle($mergedCellManpower)->applyFromArray($fontStyle);
        $sheet->getStyle($mergedCellHours)->applyFromArray($fontStyle);
    }
    // 每個 ID 的資料結束後，移動到下一區塊
    $initial_row += $block_height;
}






function dd($array)
{
    echo "<pre>";
    print_r($array);
    echo "</pre>";
}
