<?php


$member_teamwork = [

    [
        'id' => '2',
        'dispatch_company' => '華碩',
        'team_id' => '2024A',
        'team_name' => '電工C組',
        'dispatch_day' => '2',
        'construction_site' => '新竹廠',
        'manpower' => '10',
        'workinghours' => '4',
        'manpower_supported' => '10',
        'workinghours_supported' => '4',
        'foreign_manpower_supported' => '2',
        'foreign_workinghours_supported' => '2',
    ],

    [
        'id' => '2',
        'dispatch_company' => '華碩',
        'team_id' => '2024A',
        'team_name' => '電工A組',
        'dispatch_day' => '3',
        'construction_site' => '新竹廠',
        'manpower' => '10',
        'workinghours' => '4',
        'manpower_supported' => '10',
        'workinghours_supported' => '4',
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
        'manpower' => '10',
        'workinghours' => '4',
        'manpower_supported' => '10',
        'workinghours_supported' => '4',
        'foreign_manpower_supported' => '2',
        'foreign_workinghours_supported' => '2',
    ],
    [
        'id' => '3',
        'dispatch_company' => '台達電',
        'team_id' => '2024A',
        'team_name' => '電工A組',
        'dispatch_day' => '6',
        'construction_site' => '新竹廠',
        'manpower' => '10',
        'workinghours' => '4',
        'manpower_supported' => '10',
        'workinghours_supported' => '4',
        'foreign_manpower_supported' => '2',
        'foreign_workinghours_supported' => '2',
    ],
    [
        'id' => '4',
        'dispatch_company' => '聯發科',
        'team_id' => '2024A',
        'team_name' => '電工A組',
        'dispatch_day' => '8',
        'construction_site' => '新竹廠',
        'manpower' => '10',
        'workinghours' => '4',
        'manpower_supported' => '10',
        'workinghours_supported' => '4',
        'foreign_manpower_supported' => '2',
        'foreign_workinghours_supported' => '2',
    ],

];

$teamwork_map = [];

foreach ($member_teamwork as $entry) {
    $id = $entry['id'];
    $dispatch_day = $entry['dispatch_day'];

    // 確保 $teamwork_map[$id] 存在
    if (!isset($teamwork_map[$id])) {
        $teamwork_map[$id] = [];
    }

    // 確保 $teamwork_map[$id][$dispatch_day] 存在
    if (!isset($teamwork_map[$id][$dispatch_day])) {
        $teamwork_map[$id][$dispatch_day] = [];
    }

    // 加入數據
    $teamwork_map[$id][$dispatch_day][] = $entry;
}


dd($teamwork_map);
function dd($array)
{
    echo "<pre>";
    print_r($array);
    echo "</pre>";
}



