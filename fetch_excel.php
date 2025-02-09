<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;

// === Настройки ===
define('API_KEY', 'my_secret_api_key'); // Статичный API-ключ
$googleDriveFileId = '195IAlmwKs3AMGBzpPrpAegsgik0SDLt6'; // Укажите свой ID файла
$cell = 'B2'; // Укажите ячейку для чтения
$api_key = "soij091j2390ksd091k231";
$jsonFile = 'data.json'; // Файл для сохранения

/**
 * Функция получает данные из Excel-файла на Google Диске
 */
function getExcelData($googleDriveFileId, $cell)
{
    $url = "https://drive.google.com/uc?export=download&id=195IAlmwKs3AMGBzpPrpAegsgik0SDLt6";

    // Скачиваем файл
    $filePath = 'temp.xlsx';
    file_put_contents($filePath, file_get_contents($url));

    // Загружаем Excel-файл
    $spreadsheet = IOFactory::load($filePath);
    $sheet = $spreadsheet->getActiveSheet();

    // Получаем данные из указанной ячейки
    return $sheet->getCell($cell)->getValue();
}

/**
 * Функция обновляет JSON-файл с новыми данными
 */
function updateJsonFile($data, $jsonFile)
{
    $jsonData = [
        'timestamp' => date('Y-m-d H:i:s'),
        'value' => $data
    ];
    file_put_contents($jsonFile, json_encode($jsonData, JSON_PRETTY_PRINT));
}

// Бесконечный цикл для обновления данных каждые 3 минуты
while (true) {
    $data = getExcelData($googleDriveFileId, $cell);
    updateJsonFile($data, $jsonFile);

    echo "Обновлено: " . date('Y-m-d H:i:s') . " - Значение: $data\n";

    // Ждём 3 минуты (180 секунд)
    sleep(12);
}

?>