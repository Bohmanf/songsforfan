<?php
if ($_SERVER["REQUEST_METHOD"] == "POST") {
    // Hämta data från formuläret
    $artistnamn = $_POST["artistnamn"];
    $latnamn = $_POST["latnamn"];
    $status = $_POST["status"];
    $arbetsnamn = $_POST["arbetsnamn"];
    $tonart = $_POST["tonart"];

    // Skapa eller öppna Excel-filen
    require_once 'PHPExcel/IOFactory.php';
    $excel = PHPExcel_IOFactory::load('din_excel_fil.xlsx');

    // Välj ark att arbeta med (till exempel det första arket)
    $ark = $excel->getActiveSheet();

    // Hitta nästa lediga rad i Excel-filen
    $rad = $ark->getHighestRow() + 1;

    // Sätt in användardata i Excel-filen
    $ark->setCellValue('A' . $rad, $artistnamn);
    $ark->setCellValue('B' . $rad, $latnamn);
    $ark->setCellValue('C' . $rad, $status);
    $ark->setCellValue('D' . $rad, $arbetsnamn);
    $ark->setCellValue('E' . $rad, $tonart);

    // Spara Excel-filen
    $objWriter = PHPExcel_IOFactory::createWriter($excel, 'Excel2007');
    $objWriter->save('din_excel_fil.xlsx');

    // Skicka användaren tillbaka till formuläret eller en bekräftelsesida
    header("Location: index.html");
    exit();
}
?>