<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;

// === Настройки ===
$googleDriveFileId = '1poS_CPyX5vOgpTr6M9Hkt9yOeVZVG96v'; // ID файла на Google Диске
$jsonFile = 'data.json'; // Файл для сохранения
$cells = [
    'month' => 'A1',
    'number1' => ['B2', 'C2', 'D2', 'E2', 'F2', 'G2', 'H2', 'I2', 'J2', 'K2', 'L2', 'M2', 'N2', 'O2', 'P2', 'Q2', 'R2', 'S2', 'T2', 'U2', 'V2', 'W2', 'X2', 'Y2', 'Z2', 'AA2', 'AB2', 'AC2', 'AD2', 'AE2', 'AF2', 'AG2'],
    'number2' => ['B3', 'C3', 'D3', 'E3', 'F3', 'G3', 'H3', 'I3', 'J3', 'K3', 'L3', 'M3', 'N3', 'O3', 'P3', 'Q3', 'R3', 'S3', 'T3', 'U3', 'V3', 'W3', 'X3', 'Y3', 'Z3', 'AA3', 'AB3', 'AC3', 'AD3', 'AE3', 'AF3', 'AG3'],
    'number3' => ['B4', 'C4', 'D4', 'E4', 'F4', 'G4', 'H4', 'I4', 'J4', 'K4', 'L4', 'M4', 'N4', 'O4', 'P4', 'Q4', 'R4', 'S4', 'T4', 'U4', 'V4', 'W4', 'X4', 'Y4', 'Z4', 'AA4', 'AB4', 'AC4', 'AD4', 'AE4', 'AF4', 'AG4'],
    'number4' => ['B5', 'C5', 'D5', 'E5', 'F5', 'G5', 'H5', 'I5', 'J5', 'K5', 'L5', 'M5', 'N5', 'O5', 'P5', 'Q5', 'R5', 'S5', 'T5', 'U5', 'V5', 'W5', 'X5', 'Y5', 'Z5', 'AA5', 'AB5', 'AC5', 'AD5', 'AE5', 'AF5', 'AG5'],
    'number5' => ['B6', 'C6', 'D6', 'E6', 'F6', 'G6', 'H6', 'I6', 'J6', 'K6', 'L6', 'M6', 'N6', 'O6', 'P6', 'Q6', 'R6', 'S6', 'T6', 'U6', 'V6', 'W6', 'X6', 'Y6', 'Z6', 'AA6', 'AB6', 'AC6', 'AD6', 'AE6', 'AF6', 'AG6'],
    'number6' => ['B7', 'C7', 'D7', 'E7', 'F7', 'G7', 'H7', 'I7', 'J7', 'K7', 'L7', 'M7', 'N7', 'O7', 'P7', 'Q7', 'R7', 'S7', 'T7', 'U7', 'V7', 'W7', 'X7', 'Y7', 'Z7', 'AA7', 'AB7', 'AC7', 'AD7', 'AE7', 'AF7', 'AG7'],
    'number7' => ['B8', 'C8', 'D8', 'E8', 'F8', 'G8', 'H8', 'I8', 'J8', 'K8', 'L8', 'M8', 'N8', 'O8', 'P8', 'Q8', 'R8', 'S8', 'T8', 'U8', 'V8', 'W8', 'X8', 'Y8', 'Z8', 'AA8', 'AB8', 'AC8', 'AD8', 'AE8', 'AF8', 'AG8'],
    'number8' => ['B9', 'C9', 'D9', 'E9', 'F9', 'G9', 'H9', 'I9', 'J9', 'K9', 'L9', 'M9', 'N9', 'O9', 'P9', 'Q9', 'R9', 'S9', 'T9', 'U9', 'V9', 'W9', 'X9', 'Y9', 'Z9', 'AA9', 'AB9', 'AC9', 'AD9', 'AE9', 'AF9', 'AG9'],
    'number9' => ['B10', 'C10', 'D10', 'E10', 'F10', 'G10', 'H10', 'I10', 'J10', 'K10', 'L10', 'M10', 'N10', 'O10', 'P10', 'Q10', 'R10', 'S10', 'T10', 'U10', 'V10', 'W10', 'X10', 'Y10', 'Z10', 'AA10', 'AB10', 'AC10', 'AD10', 'AE10', 'AF10', 'AG10'],
    'number10' => ['B11', 'C11', 'D11', 'E11', 'F11', 'G11', 'H11', 'I11', 'J11', 'K11', 'L11', 'M11', 'N11', 'O11', 'P11', 'Q11', 'R11', 'S11', 'T11', 'U11', 'V11', 'W11', 'X11', 'Y11', 'Z11', 'AA11', 'AB11', 'AC11', 'AD11', 'AE11', 'AF11', 'AG11'],
    'number11' => ['B12', 'C12', 'D12', 'E12', 'F12', 'G12', 'H12', 'I12', 'J12', 'K12', 'L12', 'M12', 'N12', 'O12', 'P12', 'Q12', 'R12', 'S12', 'T12', 'U12', 'V12', 'W12', 'X12', 'Y12', 'Z12', 'AA12', 'AB12', 'AC12', 'AD12', 'AE12', 'AF12', 'AG12'],
    'number12' => ['B13', 'C13', 'D13', 'E13', 'F13', 'G13', 'H13', 'I13', 'J13', 'K13', 'L13', 'M13', 'N13', 'O13', 'P13', 'Q13', 'R13', 'S13', 'T13', 'U13', 'V13', 'W13', 'X13', 'Y13', 'Z13', 'AA13', 'AB13', 'AC13', 'AD13', 'AE13', 'AF13', 'AG13'],
    'number13' => ['B14', 'C14', 'D14', 'E14', 'F14', 'G14', 'H14', 'I14', 'J14', 'K14', 'L14', 'M14', 'N14', 'O14', 'P14', 'Q14', 'R14', 'S14', 'T14', 'U14', 'V14', 'W14', 'X14', 'Y14', 'Z14', 'AA14', 'AB14', 'AC14', 'AD14', 'AE14', 'AF14', 'AG14'],
    'number14' => ['B15', 'C15', 'D15', 'E15', 'F15', 'G15', 'H15', 'I15', 'J15', 'K15', 'L15', 'M15', 'N15', 'O15', 'P15', 'Q15', 'R15', 'S15', 'T15', 'U15', 'V15', 'W15', 'X15', 'Y15', 'Z15', 'AA15', 'AB15', 'AC15', 'AD15', 'AE15', 'AF15', 'AG15'],
    'number15' => ['B16', 'C16', 'D16', 'E16', 'F16', 'G16', 'H16', 'I16', 'J16', 'K16', 'L16', 'M16', 'N16', 'O16', 'P16', 'Q16', 'R16', 'S16', 'T16', 'U16', 'V16', 'W16', 'X16', 'Y16', 'Z16', 'AA16', 'AB16', 'AC16', 'AD16', 'AE16', 'AF16', 'AG16'],
    'number16' => ['B17', 'C17', 'D17', 'E17', 'F17', 'G17', 'H17', 'I17', 'J17', 'K17', 'L17', 'M17', 'N17', 'O17', 'P17', 'Q17', 'R17', 'S17', 'T17', 'U17', 'V17', 'W17', 'X17', 'Y17', 'Z17', 'AA17', 'AB17', 'AC17', 'AD17', 'AE17', 'AF17', 'AG17'],
    'number17' => ['B18', 'C18', 'D18', 'E18', 'F18', 'G18', 'H18', 'I18', 'J18', 'K18', 'L18', 'M18', 'N18', 'O18', 'P18', 'Q18', 'R18', 'S18', 'T18', 'U18', 'V18', 'W18', 'X18', 'Y18', 'Z18', 'AA18', 'AB18', 'AC18', 'AD18', 'AE18', 'AF18', 'AG18'],
    'number18' => ['B19', 'C19', 'D19', 'E19', 'F19', 'G19', 'H19', 'I19', 'J19', 'K19', 'L19', 'M19', 'N19', 'O19', 'P19', 'Q19', 'R19', 'S19', 'T19', 'U19', 'V19', 'W19', 'X19', 'Y19', 'Z19', 'AA19', 'AB19', 'AC19', 'AD19', 'AE19', 'AF19', 'AG19'],
    'number19' => ['B20', 'C20', 'D20', 'E20', 'F20', 'G20', 'H20', 'I20', 'J20', 'K20', 'L20', 'M20', 'N20', 'O20', 'P20', 'Q20', 'R20', 'S20', 'T20', 'U20', 'V20', 'W20', 'X20', 'Y20', 'Z20', 'AA20', 'AB20', 'AC20', 'AD20', 'AE20', 'AF20', 'AG20'],
    'number20' => ['B21', 'C21', 'D21', 'E21', 'F21', 'G21', 'H21', 'I21', 'J21', 'K21', 'L21', 'M21', 'N21', 'O21', 'P21', 'Q21', 'R21', 'S21', 'T21', 'U21', 'V21', 'W21', 'X21', 'Y21', 'Z21', 'AA21', 'AB21', 'AC21', 'AD21', 'AE21', 'AF21', 'AG21'],
    'number21' => ['B22', 'C22', 'D22', 'E22', 'F22', 'G22', 'H22', 'I22', 'J22', 'K22', 'L22', 'M22', 'N22', 'O22', 'P22', 'Q22', 'R22', 'S22', 'T22', 'U22', 'V22', 'W22', 'X22', 'Y22', 'Z22', 'AA22', 'AB22', 'AC22', 'AD22', 'AE22', 'AF22', 'AG22'],
    'number22' => ['B23', 'C23', 'D23', 'E23', 'F23', 'G23', 'H23', 'I23', 'J23', 'K23', 'L23', 'M23', 'N23', 'O23', 'P23', 'Q23', 'R23', 'S23', 'T23', 'U23', 'V23', 'W23', 'X23', 'Y23', 'Z23', 'AA23', 'AB23', 'AC23', 'AD23', 'AE23', 'AF23', 'AG23'],
    'number23' => ['B24', 'C24', 'D24', 'E24', 'F24', 'G24', 'H24', 'I24', 'J24', 'K24', 'L24', 'M24', 'N24', 'O24', 'P24', 'Q24', 'R24', 'S24', 'T24', 'U24', 'V24', 'W24', 'X24', 'Y24', 'Z24', 'AA24', 'AB24', 'AC24', 'AD24', 'AE24', 'AF24', 'AG24'],
    'number24' => ['B25', 'C25', 'D25', 'E25', 'F25', 'G25', 'H25', 'I25', 'J25', 'K25', 'L25', 'M25', 'N25', 'O25', 'P25', 'Q25', 'R25', 'S25', 'T25', 'U25', 'V25', 'W25', 'X25', 'Y25', 'Z25', 'AA25', 'AB25', 'AC25', 'AD25', 'AE25', 'AF25', 'AG25'],
    'number25' => ['B26', 'C26', 'D26', 'E26', 'F26', 'G26', 'H26', 'I26', 'J26', 'K26', 'L26', 'M26', 'N26', 'O26', 'P26', 'Q26', 'R26', 'S26', 'T26', 'U26', 'V26', 'W26', 'X26', 'Y26', 'Z26', 'AA26', 'AB26', 'AC26', 'AD26', 'AE26', 'AF26', 'AG26'],
    'number26' => ['B27', 'C27', 'D27', 'E27', 'F27', 'G27', 'H27', 'I27', 'J27', 'K27', 'L27', 'M27', 'N27', 'O27', 'P27', 'Q27', 'R27', 'S27', 'T27', 'U27', 'V27', 'W27', 'X27', 'Y27', 'Z27', 'AA27', 'AB27', 'AC27', 'AD27', 'AE27', 'AF27', 'AG27'],
    'number27' => ['B28', 'C28', 'D28', 'E28', 'F28', 'G28', 'H28', 'I28', 'J28', 'K28', 'L28', 'M28', 'N28', 'O28', 'P28', 'Q28', 'R28', 'S28', 'T28', 'U28', 'V28', 'W28', 'X28', 'Y28', 'Z28', 'AA28', 'AB28', 'AC28', 'AD28', 'AE28', 'AF28', 'AG28'],
    'number28' => ['B29', 'C29', 'D29', 'E29', 'F29', 'G29', 'H29', 'I29', 'J29', 'K29', 'L29', 'M29', 'N29', 'O29', 'P29', 'Q29', 'R29', 'S29', 'T29', 'U29', 'V29', 'W29', 'X29', 'Y29', 'Z29', 'AA29', 'AB29', 'AC29', 'AD29', 'AE29', 'AF29', 'AG29'],
    // 'number29' => ['B30', 'C30', 'D30', 'E30', 'F30', 'G30', 'H30', 'I30', 'J30', 'K30', 'L30', 'M30', 'N30', 'O30', 'P30', 'Q30', 'R30', 'S30', 'T30', 'U30', 'V30', 'W30', 'X30', 'Y30', 'Z30', 'AA30', 'AB30', 'AC30', 'AD30', 'AE30', 'AF30', 'AG30'],
    // 'number30' => ['B31', 'C31', 'D31', 'E31', 'F31', 'G31', 'H31', 'I31', 'J31', 'K31', 'L31', 'M31', 'N31', 'O31', 'P31', 'Q31', 'R31', 'S31', 'T31', 'U31', 'V31', 'W31', 'X31', 'Y31', 'Z31', 'AA31', 'AB31', 'AC31', 'AD31', 'AE31', 'AF31', 'AG31'],
    // 'number31' => ['B32', 'C32', 'D32', 'E32', 'F32', 'G32', 'H32', 'I32', 'J32', 'K32', 'L32', 'M32', 'N32', 'O32', 'P32', 'Q32', 'R32', 'S32', 'T32', 'U32', 'V32', 'W32', 'X32', 'Y32', 'Z32', 'AA32', 'AB32', 'AC32', 'AD32', 'AE32', 'AF32', 'AG32'],
];
function getExcelData($googleDriveFileId, $cells, $sheetIndex = 3)
{
    $url = "https://drive.google.com/uc?export=download&id=$googleDriveFileId";
    $filePath = 'temp.xlsx';
    file_put_contents($filePath, file_get_contents($url));

    $spreadsheet = IOFactory::load($filePath);

    // Получаем нужный лист по индексу
    $sheet = $spreadsheet->getSheet($sheetIndex);

    $data = [];
    foreach ($cells as $key => $cell) {
        if (is_array($cell)) {
            $data[$key] = [];
            foreach ($cell as $subCell) {
                // Получаем значение ячейки и заменяем пустое значение на 0
                $value = $sheet->getCell($subCell)->getValue();
                $data[$key][] = ($value === null || $value === '') ? "-----" : $value;
            }
        } else {
            // Получаем значение ячейки и заменяем пустое значение на 0
            $value = $sheet->getCell($cell)->getValue();
            $data[$key] = ($value === null || $value === '') ? "-----" : $value;
        }
    }
    return $data;
}



/**
 * Обновляет JSON-файл с новыми данными
 */
function updateJsonFile($data, $jsonFile)
{
    $jsonData = [
        'timestamp' => date('Y-m-d H:i:s'),
        'values' => $data
    ];
    file_put_contents($jsonFile, json_encode($jsonData, JSON_PRETTY_PRINT));
}

/**
 * API обработка запросов
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
    $data = getExcelData($googleDriveFileId, $cells);
    updateJsonFile($data, $jsonFile);
    echo json_encode(['status' => 'success', 'timestamp' => date('Y-m-d H:i:s'), 'values' => $data], JSON_PRETTY_PRINT);
    exit;
}

http_response_code(405);
echo json_encode(['error' => 'Method not allowed']);
