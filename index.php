<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;

// === Настройки ===
$googleDriveFileId = '1poS_CPyX5vOgpTr6M9Hkt9yOeVZVG96v'; // ID файла на Google Диске
$cells = [
    'month' => 'A1',
    'number1' => ['yurt1' => 'B2', 'yurt2' => 'C2', 'yurt3' => 'D2', 'yurt4' => 'E2', 'yurt5' => 'F2', 'yurt6' => 'G2', 'yurt7' => 'H2', 'yurt8' => 'I2', 'yurt9' => 'J2', 'yurt10' => 'K2', 'yurt11' => 'L2', 'yurt12' => 'M2', 'yurt13' => 'N2', 'yurt14' => 'O2', 'yurt15' => 'P2', 'yurt16' => 'Q2', 'yurt17' => 'R2', 'yurt18' => 'S2', 'yurt19' => 'T2', 'yurt20' => 'U2', 'yurt21' => 'V2', 'yurt22' => 'W2', 'yurt23' => 'X2', 'yurt24' => 'Y2', 'yurt25' => 'Z2', 'yurt26' => 'AA2', 'yurt27' => 'AB2', 'yurt28' => 'AC2', 'yurt29' => 'AD2', 'yurt30' => 'AE2', 'yurt31' => 'AF2', 'yurt32' => 'AG2'],
];
$jsonFile = 'data.json'; // Файл для сохранения

/**
 * Функция получает данные из Excel-файла на Google Диске
 */
function getExcelData($googleDriveFileId, $cells)
{
    $url = "https://drive.google.com/uc?export=download&id=$googleDriveFileId";

    // Скачиваем файл
    $filePath = 'temp.xlsx';
    file_put_contents($filePath, file_get_contents($url));

    // Загружаем Excel-файл
    $spreadsheet = IOFactory::load($filePath);
    $sheet = $spreadsheet->getActiveSheet();

    // Получаем данные из указанных ячеек
    $data = [];
    foreach ($cells as $key => $cell) {
        if (is_array($cell)) {
            // Если несколько ячеек, получаем массив значений
            $data[$key] = [];
            foreach ($cell as $subCell) {
                $data[$key][] = $sheet->getCell($subCell)->getValue();
            }
        } else {
            // Если одна ячейка, получаем одно значение
            $data[$key] = $sheet->getCell($cell)->getValue();
        }
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
        'month' => $data['month']
    ];

    // Убираем "month" из основного массива данных, чтобы не дублировать
    unset($data['month']);

    // Добавляем остальные данные
    $jsonData += $data;

    file_put_contents($jsonFile, json_encode($jsonData, JSON_PRETTY_PRINT | JSON_UNESCAPED_UNICODE));
}


/**
 * Проверка времени последнего обновления
 */
if (file_exists($jsonFile)) {
    $jsonData = json_decode(file_get_contents($jsonFile), true);
    $lastUpdated = strtotime($jsonData['timestamp']);
    $currentTime = time();

    // Если прошло менее 60 секунд - не обновляем
    if (($currentTime - $lastUpdated) < 60) {
        header('Content-Type: application/json');
        echo json_encode($jsonData, JSON_PRETTY_PRINT);
        exit;
    }
}

// Если обновление требуется - получаем новые данные
$data = getExcelData($googleDriveFileId, $cells);
updateJsonFile($data, $jsonFile);

// Отправляем свежие данные
header('Content-Type: application/json');
echo json_encode(['timestamp' => date('Y-m-d H:i:s'), 'values' => $data], JSON_PRETTY_PRINT);
exit;

?>