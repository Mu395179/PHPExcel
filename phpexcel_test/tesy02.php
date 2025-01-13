<?php
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

dd($member_teamwork);
$teamwork_map = [];
foreach ($member_teamwork as $entry) {
    $teamwork_map[$entry['id']][$entry['dispatch_day']] = $entry;
    dd($teamwork_map);
}

function dd($array)
{
    echo "<pre>";
    print_r($array);
    echo "</pre>";
}