<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;

// === Настройки ===
$googleDriveFileId = '1poS_CPyX5vOgpTr6M9Hkt9yOeVZVG96v'; // ID файла на Google Диске
$cells = [
    'month' => 'A1',
    'numbers' => []
];

for ($row = 2; $row <= 32; $row++) {
    $rowIndex = $row - 1; // Индекс строки (1, 2, 3 и т. д.)
    $cells['numbers']["number$rowIndex"] = []; // Инициализируем массив

    $yurtIndex = 1;
    for ($col = 'B'; $col !== 'AH'; $col++) {
        $cells['numbers']["number$rowIndex"]["yurt$yurtIndex"] = $col . $row;
        $yurtIndex++;
    }
}

$json = json_encode($cells, JSON_PRETTY_PRINT | JSON_UNESCAPED_UNICODE);
echo $json;


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
