<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;

// === Настройки ===
define('API_KEY', 'your_static_api_key_here'); // Статичный API-ключ
$googleDriveFileId = '1XyDAvetu-aqH5pBh1Yn0up7khtt0crdL'; // ID файла на Google Диске
$cells = [
    'A1', 'A2', 'A3', 'A4', 'A5', 'A6', 'A7', 'A8', 'A9', 'A10', 'A11', 'A12', 'A13', 'A14', 'A15', 'A16',
    'B1', 'B2', 'B3', 'B4', 'B5', 'B6', 'B7', 'B8', 'B9', 'B10', 'B11', 'B12', 'B13', 'B14', 'B15', 'B16',
    'C1', 'C2', 'C3', 'C4', 'C5', 'C6', 'C7', 'C8', 'C9', 'C10', 'C11', 'C12', 'C13', 'C14', 'C15', 'C16',
    'D1', 'D2', 'D3', 'D4', 'D5', 'D6', 'D7', 'D8', 'D9', 'D10', 'D11', 'D12', 'D13', 'D14', 'D15', 'D16'
];
$jsonFile = 'data.json'; // Файл для сохранения

/**
 * Функция получает данные из Excel-файла на Google Диске
 */
function getExcelData($googleDriveFileId, $cells)
{
    $url = "https://drive.google.com/uc?export=download&id=1XyDAvetu-aqH5pBh1Yn0up7khtt0crdL";
    
    // Скачиваем файл
    $filePath = 'temp.xlsx';
    file_put_contents($filePath, file_get_contents($url));

    // Загружаем Excel-файл
    $spreadsheet = IOFactory::load($filePath);
    $sheet = $spreadsheet->getActiveSheet();

    // Получаем данные из указанных ячеек
    $data = [];
    foreach ($cells as $cell) {
        $data[$cell] = $sheet->getCell($cell)->getValue();
    }
    return $data;
}

/**
 * Функция обновляет JSON-файл с новыми данными
 */
function updateJsonFile($data, $jsonFile)
{
    $jsonData = [
        'timestamp' => date('Y-m-d H:i:s'),
        'values' => $data
    ];
    file_put_contents($jsonFile, json_encode($jsonData, JSON_PRETTY_PRINT));
}

// Проверяем API-ключ
if (!isset($_GET['api_key']) || $_GET['api_key'] !== API_KEY) {
    http_response_code(403);
    echo json_encode(['error' => 'Invalid API key']);
    exit;
}

// Проверяем, был ли запрос на обновление
if (isset($_GET['update']) && $_GET['update'] === 'true') {
    $data = getExcelData($googleDriveFileId, $cells);
    updateJsonFile($data, $jsonFile);
    
    header('Content-Type: application/json');
    echo json_encode(['timestamp' => date('Y-m-d H:i:s'), 'values' => $data], JSON_PRETTY_PRINT);
    exit;
}

// Выдача данных из JSON-файла
if (file_exists($jsonFile)) {
    header('Content-Type: application/json');
    echo file_get_contents($jsonFile);
    exit;
} else {
    http_response_code(404);
    echo json_encode(['error' => 'Data not found']);
    exit;
}

?>
