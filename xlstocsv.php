<?php

use PayoXlsToCsv\PayoXlsToCsv;

require_once "vendor/autoload.php";

$xls_dir = __DIR__ . '/files/Xls/';
$csv_dir = '/files/Csv/';
$files = array_diff(scandir($xls_dir), array('.', '..', 'processed'));
foreach ($files as $xlsFile) {
    $x = new PayoXlsToCsv(__DIR__ . $csv_dir);
    $x->convert($xls_dir . $xlsFile);
    rename($xls_dir . $xlsFile, $xls_dir . "/processed/" . basename($xlsFile));
    break;
}