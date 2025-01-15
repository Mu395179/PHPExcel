<?php

//error_reporting(E_ALL); 
//ini_set('display_errors', '1');

session_start();

$memberID = $_SESSION['memberID'];
$powerkey = $_SESSION['powerkey'];

/**
 * PHPExcel
 *
 * Copyright (C) 2006 - 2012 PHPExcel
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 *
 * @category   PHPExcel
 * @package    PHPExcel
 * @copyright  Copyright (c) 2006 - 2012 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    1.7.8, 2012-10-12
 */

/** Error reporting */
/*
error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
date_default_timezone_set("Asia/Taipei");
*/
ini_set('display_errors', FALSE);
ini_set('display_startup_errors', FALSE);
date_default_timezone_set("Asia/Taipei");




//計算請假時間
function calculateLeaveHours($startTime, $endTime)
{
	// 設定工作時間的起點與終點
	$workStartTime = new DateTime('08:00');
	$workEndTime = new DateTime('17:00');

	// 中午休息時間
	$lunchStartTime = new DateTime('12:00');
	$lunchEndTime = new DateTime('13:00');

	// 將請假開始時間與結束時間轉為 DateTime 物件
	$start = new DateTime($startTime);
	$end = new DateTime($endTime);

	// 檢查是否在工作時間範圍內
	if ($start < $workStartTime)
		$start = $workStartTime;
	if ($end > $workEndTime)
		$end = $workEndTime;

	// 計算總請假時數（包含中午時間）
	$interval = $start->diff($end);
	$leaveHours = $interval->h + ($interval->i / 60);

	// 檢查是否跨越中午休息時間，並扣除 1 小時
	if ($start < $lunchEndTime && $end > $lunchStartTime) {
		$leaveHours -= 1; // 扣除 1 小時中午休息時間
	}

	// 將請假時間以半小時為單位計算
	$leaveHours = ceil($leaveHours * 2) / 2;

	return $leaveHours;
}




//載入公用函數
@include_once '/website/include/pub_function.php';


$site_db = "eshop";
$web_id = "sales.eshop";

$team_id = $_GET['team_id'];
$team_id2 = $_GET['team_id2'];
$team_construction_id = $_GET['team_construction_id'];

$start_date = $_GET['start_date'];
$end_date = $_GET['end_date'];

$team_row = getkeyvalue2($site_db . '_info', 'team', "team_id = '$team_id'", 'team_name');
$team_name = $team_row['team_name'];


/*
//檢查是否為管理員及進階會員
$super_admin = "N";
$super_advanced = "N";
$mem_row = getkeyvalue2('memberinfo','member',"member_no = '$memberID'",'admin,advanced,checked,luck,admin_readonly,advanced_readonly');
$super_admin = $mem_row['admin'];
$super_advanced = $mem_row['advanced'];
*/

@include_once("/website/class/" . $site_db . "_info_class.php");


if (PHP_SAPI == 'cli')
	die('This programe should only be run from a Web Browser');

/** Include PHPExcel */
require_once '/website/os/PHPExcel-1.8.1/Classes/PHPExcel.php';


// Create new PHPExcel object
$objPHPExcel = new PHPExcel();

// Set document properties
$objPHPExcel->getProperties()->setCreator("PowerSales")
	->setLastModifiedBy("PowerSales")
	->setTitle("Office 2007 XLSX Document")
	->setSubject("Office 2007 XLSX Document")
	->setDescription("The document for Office 2007 XLSX, generated using PHP classes.")
	->setKeywords("office 2007 openxml php")
	->setCategory($team_name . "_" . $start_date . "~" . $end_date . "_團隊支援表");



// 設定起始日期和結束日期
$date1 = new DateTime($start_date);
$date2 = new DateTime($end_date);




$mDB = "";
$mDB = new MywebDB();

$mDB2 = "";
$mDB2 = new MywebDB();


//取得工單資料
if (!empty($team_construction_id)) {
	if (!empty($team_id2)) {
		$Qry = "SELECT a.dispatch_id,YEAR(a.dispatch_date) AS dispatch_year,MONTH(a.dispatch_date) AS dispatch_month,DAY(a.dispatch_date) AS dispatch_day,b.employee_id,c.employee_name,a.team_id,d.team_name 
		,b.team_id as team_id2,e.team_name as team_name2,b.team_construction_id,f.construction_site
		FROM dispatch a
		LEFT JOIN dispatch_member b ON b.dispatch_id = a.dispatch_id
		LEFT JOIN employee c ON c.employee_id = b.employee_id
		LEFT JOIN team d ON d.team_id = a.team_id
		LEFT JOIN team e ON e.team_id = b.team_id
		LEFT JOIN construction f ON f.construction_id = b.team_construction_id
		WHERE a.dispatch_date >= '$start_date' AND a.dispatch_date <= '$end_date' AND a.ConfirmSending = 'Y' 
		AND a.team_id = '$team_id' AND b.team_id is not null AND b.team_id <> '' 
		AND b.team_id = '$team_id2' AND b.team_construction_id = '$team_construction_id'
		GROUP BY b.employee_id
		ORDER BY c.team_id,b.employee_id";
	} else {
		$Qry = "SELECT a.dispatch_id,YEAR(a.dispatch_date) AS dispatch_year,MONTH(a.dispatch_date) AS dispatch_month,DAY(a.dispatch_date) AS dispatch_day,b.employee_id,c.employee_name,a.team_id,d.team_name 
		,b.team_id as team_id2,e.team_name as team_name2,b.team_construction_id,f.construction_site
		FROM dispatch a
		LEFT JOIN dispatch_member b ON b.dispatch_id = a.dispatch_id
		LEFT JOIN employee c ON c.employee_id = b.employee_id
		LEFT JOIN team d ON d.team_id = a.team_id
		LEFT JOIN team e ON e.team_id = b.team_id
		LEFT JOIN construction f ON f.construction_id = b.team_construction_id
		WHERE a.dispatch_date >= '$start_date' AND a.dispatch_date <= '$end_date' AND a.ConfirmSending = 'Y' 
		AND a.team_id = '$team_id' AND b.team_id is not null AND b.team_id <> '' 
		AND b.team_construction_id = '$team_construction_id'
		GROUP BY b.employee_id
		ORDER BY c.team_id,b.employee_id";
	}
} else {
	if (!empty($team_id2)) {
		$Qry = "SELECT a.dispatch_id,YEAR(a.dispatch_date) AS dispatch_year,MONTH(a.dispatch_date) AS dispatch_month,DAY(a.dispatch_date) AS dispatch_day,b.employee_id,c.employee_name,a.team_id,d.team_name 
		,b.team_id as team_id2,e.team_name as team_name2,b.team_construction_id,f.construction_site
		FROM dispatch a
		LEFT JOIN dispatch_member b ON b.dispatch_id = a.dispatch_id
		LEFT JOIN employee c ON c.employee_id = b.employee_id
		LEFT JOIN team d ON d.team_id = a.team_id
		LEFT JOIN team e ON e.team_id = b.team_id
		LEFT JOIN construction f ON f.construction_id = b.team_construction_id
		WHERE a.dispatch_date >= '$start_date' AND a.dispatch_date <= '$end_date' AND a.ConfirmSending = 'Y' 
		AND a.team_id = '$team_id' AND b.team_id is not null AND b.team_id <> '' 
		AND b.team_id = '$team_id2'
		GROUP BY b.employee_id
		ORDER BY c.team_id,b.employee_id";
	} else {
		$Qry = "SELECT a.dispatch_id,YEAR(a.dispatch_date) AS dispatch_year,MONTH(a.dispatch_date) AS dispatch_month,DAY(a.dispatch_date) AS dispatch_day,b.employee_id,c.employee_name,a.team_id,d.team_name 
		,b.team_id as team_id2,e.team_name as team_name2,b.team_construction_id,f.construction_site
		FROM dispatch a
		LEFT JOIN dispatch_member b ON b.dispatch_id = a.dispatch_id
		LEFT JOIN employee c ON c.employee_id = b.employee_id
		LEFT JOIN team d ON d.team_id = a.team_id
		LEFT JOIN team e ON e.team_id = b.team_id
		LEFT JOIN construction f ON f.construction_id = b.team_construction_id
		WHERE a.dispatch_date >= '$start_date' AND a.dispatch_date <= '$end_date' AND a.ConfirmSending = 'Y' 
		AND a.team_id = '$team_id' AND b.team_id is not null AND b.team_id <> '' 
		GROUP BY b.employee_id
		ORDER BY c.team_id,b.employee_id";
	}
}


$mDB->query($Qry);

$total = $mDB->rowCount();

$line = 1;

if ($total > 0) {



	$seq = 0;
	while ($row = $mDB->fetchRow(2)) {
		$dispatch_id = $row['dispatch_id'];
		$dispatch_year = $row['dispatch_year'];
		$dispatch_month = $row['dispatch_month'];
		$dispatch_day = $row['dispatch_day'];
		$employee_id = $row['employee_id'];
		$employee_name = $row['employee_name'];
		$team_id = $row['team_id'];
		$team_name = $row['team_name'];

		$team_id2 = $row['team_id2'];
		$team_name2 = $row['team_name2'];

		$team_construction_id = $row['team_construction_id'];
		$construction_site = $row['construction_site'];

		if (!empty($employee_id)) {


			//再取得各員工的資料
			$Qry2 = "SELECT a.dispatch_id,a.dispatch_date,YEAR(a.dispatch_date) AS dispatch_year,MONTH(a.dispatch_date) AS dispatch_month,DAY(a.dispatch_date) AS dispatch_day
				,b.employee_id,b.manpower,b.attendance_day,b.attendance_status,b.transition,b.transition_start,b.transition_end
				,c.team_name,d.team_name as transition_team_name
				FROM dispatch a
				LEFT JOIN dispatch_member b ON b.dispatch_id = a.dispatch_id
				LEFT JOIN team c ON c.team_id = b.team_id
				LEFT JOIN team d ON d.team_id = b.transition_team_id
				WHERE a.dispatch_date >= '$start_date' AND a.dispatch_date <= '$end_date' AND a.ConfirmSending = 'Y' AND b.employee_id = '$employee_id'
				ORDER BY b.employee_id";

			$mDB2->query($Qry2);
			$SUMMARY = 0;
			if ($mDB2->rowCount() > 0) {

			

				while ($row2 = $mDB2->fetchRow(2)) {

					

					
				}
			}
		}
	}
}


$mDB2->remove();
$mDB->remove();


// Rename worksheet
$objPHPExcel->getActiveSheet()->setTitle("團隊支援表");


$xlsx_filename = $team_name . "_" . $start_date . "~" . $end_date . "_團隊支援表.xls";


// Redirect output to a client’s web browser (Excel5)
header('Content-Type: application/vnd.ms-excel');
header('Content-Disposition: attachment;filename=' . $xlsx_filename);
header('Cache-Control: max-age=0');
// If you're serving to IE 9, then the following may be needed
header('Cache-Control: max-age=1');

// If you're serving to IE over SSL, then the following may be needed
header('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT'); // always modified
header('Cache-Control: cache, must-revalidate'); // HTTP/1.1
header('Pragma: public'); // HTTP/1.0

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save('php://output');
exit;




?>