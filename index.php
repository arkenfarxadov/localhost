<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;

// === Настройки ===
$googleDriveFileId = '1poS_CPyX5vOgpTr6M9Hkt9yOeVZVG96v'; // ID файла на Google Диске
$jsonFile = 'data.json'; // Файл для сохранения

// Определение ячеек, которые нужно получить
$cells = [
    'month' => 'A1',
    'numbers' => []
];

// Автоматическое заполнение номеров строк и колонок
for ($row = 2; $row <= 32; $row++) {
    for ($col = 'B'; $col !== 'AH'; $col++) {
        $cells['numbers'][$row - 1][] = $col . $row;
    }
}

/**
 * Функция получает данные из Excel-файла на Google Диске
 */
function getExcelData($googleDriveFileId, $cells)
{
    $url = "https://drive.google.com/uc?export=download&id=$googleDriveFileId";
    $filePath = 'temp.xlsx';
    file_put_contents($filePath, file_get_contents($url));
    
    $spreadsheet = IOFactory::load($filePath);
    $sheet = $spreadsheet->getActiveSheet();
    
    $data = ['month' => $sheet->getCell($cells['month'])->getValue(), 'numbers' => []];
    
    foreach ($cells['numbers'] as $key => $cellGroup) {
        foreach ($cellGroup as $cell) {
            $data['numbers'][$key][] = $sheet->getCell($cell)->getValue();
        }
    }
    
    return $data;
}

/**
 * Функция обновляет или создаёт JSON-файл с новыми данными
 */
function updateJsonFile($data, $jsonFile)
{
    $jsonData = [
        'timestamp' => date('Y-m-d H:i:s'),
        'month' => $data['month'],
        'numbers' => $data['numbers']
    ];
    
    file_put_contents($jsonFile, json_encode($jsonData, JSON_PRETTY_PRINT | JSON_UNESCAPED_UNICODE));
}

/**
 * Проверка файла и времени последнего обновления
 */
if (file_exists($jsonFile)) {
    $jsonData = json_decode(file_get_contents($jsonFile), true);
    $lastUpdated = strtotime($jsonData['timestamp'] ?? '');
    $currentTime = time();
    
    if ($lastUpdated && ($currentTime - $lastUpdated) < 60) {
        header('Content-Type: application/json');
        echo json_encode($jsonData, JSON_PRETTY_PRINT | JSON_UNESCAPED_UNICODE);
        exit;
    }
}

// Если обновление требуется или файла нет - получаем новые данные
$data = getExcelData($googleDriveFileId, $cells);
updateJsonFile($data, $jsonFile);

// Отправляем свежие данные
header('Content-Type: application/json');
echo json_encode(['timestamp' => date('Y-m-d H:i:s'), 'values' => $data], JSON_PRETTY_PRINT | JSON_UNESCAPED_UNICODE);
exit;

?>
