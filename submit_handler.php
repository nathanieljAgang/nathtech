<?php
// submit_handler.php
require 'vendor/autoload.php'; // Ensure PhpSpreadsheet is installed via composer

use PhpOffice\PhpSpreadsheet\IOFactory;

if ($_SERVER["REQUEST_METHOD"] == "POST") {
    // 1. Capture incoming form fields securely
    $fullname   = htmlspecialchars($_POST['fullname'] ?? '');
    $dob        = htmlspecialchars($_POST['dob'] ?? '');
    $gender     = htmlspecialchars($_POST['gender'] ?? '');
    $phone      = htmlspecialchars($_POST['phone'] ?? '');
    $email      = htmlspecialchars($_POST['email'] ?? '');
    $lga        = htmlspecialchars($_POST['lga'] ?? '');
    $state      = htmlspecialchars($_POST['state'] ?? 'Kaduna');
    $primary    = htmlspecialchars($_POST['primary_track'] ?? '');
    
    // Checkboxes
    $freelance  = isset($_POST['path_freelance']) ? 'Yes' : 'No';
    $business   = isset($_POST['path_business']) ? 'Yes' : 'No';
    $startup    = isset($_POST['path_startup']) ? 'Yes' : 'No';
    
    // Infrastructure
    $laptop     = htmlspecialchars($_POST['laptop'] ?? 'No');
    $literacy   = htmlspecialchars($_POST['literacy'] ?? 'No');
    $purpose    = htmlspecialchars($_POST['purpose'] ?? '');

    // Metadata
    $reg_date   = date("Y-m-d");
    $app_id     = "TAG-" . date("Y") . "-" . rand(100, 999);

    // 2. Load the target Excel workbook
    $inputFileName = 'tagham_foundation_enrollment_data.xlsx';
    
    try {
        $spreadsheet = IOFactory::load($inputFileName);
        $sheet = $spreadsheet->getSheetByName('Enrollment Data'); // Targeted sheet
        
        // Find the next available empty row
        $highestRow = $sheet->getHighestRow();
        $nextRow = $highestRow + 1;

        // 3. Append the data fields sequentially to match column mapping
        $sheet->setCellValue('A' . $nextRow, $app_id);
        $sheet->setCellValue('B' . $nextRow, $reg_date);
        $sheet->setCellValue('C' . $nextRow, $fullname);
        $sheet->setCellValue('D' . $nextRow, $dob);
        $sheet->setCellValue('E' . $nextRow, $gender);
        $sheet->setCellValue('F' . $nextRow, $phone);
        $sheet->setCellValue('G' . $nextRow, $email);
        $sheet->setCellValue('H' . $nextRow, $lga);
        $sheet->setCellValue('I' . $nextRow, $state);
        $sheet->setCellValue('J' . $nextRow, $primary);
        $sheet->setCellValue('K' . $nextRow, 'None Chosen'); // Secondary track placeholder
        $sheet->setCellValue('L' . $nextRow, $freelance);
        $sheet->setCellValue('M' . $nextRow, $business);
        $sheet->setCellValue('N' . $nextRow, $startup);
        $sheet->setCellValue('O' . $nextRow, 'General Enterprise');
        $sheet->setCellValue('P' . $nextRow, 'Unemployed Youth');
        $sheet->setCellValue('Q' . $nextRow, $laptop);
        $sheet->setCellValue('R' . $nextRow, $literacy);
        $sheet->setCellValue('S' . $nextRow, $purpose);

        // Optional: Style the newly appended row to match the regular font scheme
        $sheet->getStyle("A$nextRow:S$nextRow")->getFont()->setName('Segoe UI')->setSize(11);

        // 4. Save modifications back to the spreadsheet file
        $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
        $writer->save($inputFileName);

        // Redirect student to a professional success screen
        echo "<script>alert('Application Submitted Successfully! Your ID is: $app_id'); window.location.href='index.html';</script>";

    } catch (\PhpOffice\PhpSpreadsheet\Reader\Exception $e) {
        die('Error loading file: ' . $e->getMessage());
    }
}
?>

