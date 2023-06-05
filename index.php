<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\{Alignment};
use PhpOffice\PhpSpreadsheet\Writer\Exception;

$style = [
    'font' => [
        'name' => 'Arial',
        'bold' => true,
    ],
    'alignment' => [
        'horizontal' => Alignment::HORIZONTAL_CENTER,
        'vertical' => Alignment::VERTICAL_CENTER,
        'wrapText' => false,
    ]
];

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->getStyle('A1')->applyFromArray($style);
$sheet->getStyle('B1')->applyFromArray($style);
$sheet->getStyle('C1')->applyFromArray($style);
$sheet->setCellValue('A1', 'NAME');
$sheet->setCellValue('B1', 'SURNAME');
$sheet->setCellValue('C1', 'EMAIL');
$sheet->setCellValue('A2', 'Muhsin');
$sheet->setCellValue('B2', 'Shokirov');
$sheet->setCellValue('C2', 'muxsin.shokirov@gmail.com');

try {
    $writer = new Xlsx($spreadsheet);
    $writer->save('Лебер.xlsx');
} catch (Exception $e) {
    echo $e->getMessage();
}
