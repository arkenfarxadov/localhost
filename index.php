<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;

// === Настройки ===
$googleDriveFileId = '1poS_CPyX5vOgpTr6M9Hkt9yOeVZVG96v'; // ID файла на Google Диске
$jsonFile = 'data.json'; // Файл для сохранения
$defaultSheetName = 'Август'; // Имя листа по умолчанию

// Ячейки, которые нужно считать
$cells = [
    'month' => 'A1',
    'number1' => ['B2', 'C2', 'D2', 'E2', 'F2', 'G2', 'H2', 'I2', 'J2', 'K2', 'L2', 'M2', 'N2', 'O2', 'P2', 'Q2', 'R2', 'S2', 'T2', 'U2', 'V2', 'W2', 'X2', 'Y2', 'Z2', 'AA2', 'AB2', 'AC2', 'AD2', 'AE2', 'AF2', 'AG2'],
    'number2' => ['B3', 'C3', 'D3', 'E3', 'F3', 'G3', 'H3', 'I3', 'J3', 'K3', 'L3', 'M3', 'N3', 'O3', 'P3', 'Q3', 'R3', 'S3', 'T3', 'U3', 'V3', 'W3', 'X3', 'Y3', 'Z3', 'AA3', 'AB3', 'AC3', 'AD3', 'AE3', 'AF3', 'AG3'],
    // Добавь остальные ячейки, если нужно
];

/**
 * Получает данные из Excel-файла.
 */
function getExcelData($googleDriveFileId, $cells, $sheetName)
{
    $url = "https://drive.google.com/uc?export=download&id=$googleDriveFileId";
    $filePath = 'temp.xlsx';
    file_put_contents($filePath, file_get_contents($url));

    $spreadsheet = IOFactory::load($filePath);

    // Получаем список всех листов
    $allSheetNames = $spreadsheet->getSheetNames();

    // Проверяем, существует ли указанный лист
    if (!in_array($sheetName, $allSheetNames)) {
        die(json_encode(['error' => "Лист '$sheetName' не найден. Доступные листы: " . implode(', ', $allSheetNames)], JSON_PRETTY_PRINT));
    }

    // Получаем нужный лист по имени
    $sheet = $spreadsheet->getSheetByName($sheetName);

    $data = [];
    foreach ($cells as $key => $cell) {
        if (is_array($cell)) {
            $data[$key] = [];
            foreach ($cell as $subCell) {
                $value = $sheet->getCell($subCell)->getValue();
                $data[$key][] = ($value === null || $value === '') ? "-----" : $value;
            }
        } else {
            $value = $sheet->getCell($cell)->getValue();
            $data[$key] = ($value === null || $value === '') ? "-----" : $value;
        }
    }

    return $data;
}

/**
 * Обновляет JSON-файл с новыми данными.
 */
function updateJsonFile($data, $jsonFile)
{
    $jsonData = [
        'timestamp' => date('Y-m-d H:i:s'),
        'values' => $data
    ];

    $result = file_put_contents($jsonFile, json_encode($jsonData, JSON_PRETTY_PRINT | JSON_UNESCAPED_UNICODE));

    if ($result === false) {
        die(json_encode(['error' => 'Ошибка записи в файл'], JSON_PRETTY_PRINT));
    }
}

/**
 * API обработка запросов.
 */
header('Content-Type: application/json');

if ($_SERVER['REQUEST_METHOD'] === 'GET') {
    if (file_exists($jsonFile)) {
        echo file_get_contents($jsonFile);
    } else {
        echo json_encode(['error' => 'No data available']);
    }
    exit;
}

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    // Получаем имя листа из запроса или используем значение по умолчанию
    $postData = json_decode(file_get_contents("php://input"), true);
    $sheetName = $postData['sheetName'] ?? $defaultSheetName;

    $data = getExcelData($googleDriveFileId, $cells, $sheetName);
    updateJsonFile($data, $jsonFile);

    echo json_encode([
        'status' => 'success',
        'timestamp' => date('Y-m-d H:i:s'),
        'sheet' => $sheetName,
        'values' => $data
    ], JSON_PRETTY_PRINT);
    exit;
}

http_response_code(405);
echo json_encode(['error' => 'Method not allowed']);
