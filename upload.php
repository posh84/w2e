<?php

// Tắt Notice warnings
error_reporting(E_ALL & ~E_NOTICE);
ini_set('display_errors', 1);
ini_set('display_startup_errors', 1);

// Include file chứa các hàm tiện ích
//require_once 'includes/functions.php';
require 'vendor/autoload.php';

// Tạo và cấu hình thư mục temp
$tempDir = __DIR__ . '/temp/';
$imagesTempDir = $tempDir . 'images/';

if (!is_dir($tempDir)) {
    mkdir($tempDir, 0777, true);
}
if (!is_dir($imagesTempDir)) {
    mkdir($imagesTempDir, 0777, true);
}

// Thiết lập temp directory
ini_set('sys_temp_dir', $tempDir);
putenv('TMPDIR=' . $tempDir);

use PhpOffice\PhpWord\IOFactory;
use PhpOffice\PhpWord\Settings;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

// Cấu hình PhpWord
Settings::setTempDir($imagesTempDir);

// Hàm dọn dẹp temp files
function cleanupTempFiles($tempDir, $maxAge = 3600) {
    if (!is_dir($tempDir)) return;
    
    $iterator = new RecursiveIteratorIterator(
        new RecursiveDirectoryIterator($tempDir, RecursiveDirectoryIterator::SKIP_DOTS),
        RecursiveIteratorIterator::CHILD_FIRST
    );
    
    $now = time();
    foreach ($iterator as $file) {
        if ($file->isFile() && ($now - $file->getMTime()) >= $maxAge) {
            unlink($file->getRealPath());
        }
    }
}

// Thông tin kết nối database
$db_host = 'mysql.eqlab.vn';
$db_user = 'root';
$db_pass = 'EugIDhwU1vwwJUz2+mgO8kU320a8imyT';
$db_name = 'eqlabvn_lab';

// Kết nối database
try {
    $pdo = new PDO("mysql:host=$db_host;dbname=$db_name;charset=utf8mb4", $db_user, $db_pass);
    $pdo->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
} catch (PDOException $e) {
    die("Lỗi kết nối database: " . $e->getMessage());
}

// Hàm ánh xạ cột
function mapColumnToField($columnTitle) {
    $columnTitle = strtolower(trim($columnTitle));
    $columnTitle = preg_replace('/[^\p{L}\p{N}\s]/u', '', $columnTitle);
    
    $mappings = [
        'thông số phân tích' => 'parameter',
        'thong so phan tich' => 'parameter',
        'parameter' => 'parameter',
        'tham số' => 'parameter',
        'chỉ tiêu' => 'parameter',
        'chi tieu' => 'parameter',
        
        'kết quả' => 'result',
        'ket qua' => 'result',
        'result' => 'result',
        'giá trị' => 'result',
        'gia tri' => 'result',
        
        'đơn vị' => 'unit',
        'don vi' => 'unit',
        'unit' => 'unit',
        
        'phương pháp' => 'method',
        'phuong phap' => 'method',
        'method' => 'method',
        'phương pháp thử' => 'method',
        
        'giới hạn phát hiện' => 'detection_limit',
        'gioi han phat hien' => 'detection_limit',
        'detection limit' => 'detection_limit',
        
        'quy chuẩn' => 'standard',
        'quy chuan' => 'standard',
        'standard' => 'standard',
        'tiêu chuẩn' => 'standard'
    ];
    
    if (isset($mappings[$columnTitle])) {
        return $mappings[$columnTitle];
    }
    
    foreach ($mappings as $pattern => $fieldName) {
        if (strpos($columnTitle, $pattern) !== false || strpos($pattern, $columnTitle) !== false) {
            return $fieldName;
        }
    }
    
    return null;
}

// Hàm trích xuất tất cả các ID mẫu từ file Word
function extractAllSampleIds($phpWord) {
    $sampleIds = [];
    
    foreach ($phpWord->getSections() as $section) {
        foreach ($section->getElements() as $element) {
            if ($element instanceof \PhpOffice\PhpWord\Element\TextRun) {
                $text = extractTextFromTextRun($element);
                
                if (is_string($text) && !empty($text)) {
                    // Pattern 1: "Số: IER - Nxxxx.xxxx"
                    if (preg_match('/Số:\s*IER\s*-\s*(N[0-9]+\.[0-9]+(?:\/[0-9]+)?)/i', $text, $matches)) {
                        $sampleIds[] = trim($matches[1]);
                    }
                    // Pattern 2: Direct "Nxxxx.xxxx"
                    elseif (preg_match('/\b(N[0-9]+\.[0-9]+(?:\/[0-9]+)?)\b/i', $text, $matches)) {
                        $sampleIds[] = trim($matches[1]);
                    }
                    // Pattern 3: "Mẫu số/Số mẫu/Mã mẫu: Nxxxx.xxxx"
                    elseif (preg_match('/(?:Mẫu số|Số mẫu|Mã mẫu)[:.\s]+(N[0-9]+\.[0-9]+(?:\/[0-9]+)?)/i', $text, $matches)) {
                        $sampleIds[] = trim($matches[1]);
                    }
                }
            } elseif (method_exists($element, 'getText')) {
                $text = $element->getText();
                
                if (is_string($text) && !empty($text)) {
                    if (preg_match('/Số:\s*IER\s*-\s*(N[0-9]+\.[0-9]+(?:\/[0-9]+)?)/i', $text, $matches)) {
                        $sampleIds[] = trim($matches[1]);
                    }
                    elseif (preg_match('/\b(N[0-9]+\.[0-9]+(?:\/[0-9]+)?)\b/i', $text, $matches)) {
                        $sampleIds[] = trim($matches[1]);
                    }
                    elseif (preg_match('/(?:Mẫu số|Số mẫu|Mã mẫu)[:.\s]+(N[0-9]+\.[0-9]+(?:\/[0-9]+)?)/i', $text, $matches)) {
                        $sampleIds[] = trim($matches[1]);
                    }
                }
            }
        }
    }
    
    return array_unique($sampleIds);
}

// Các hàm hỗ trợ khác giữ nguyên
function sanitizeFilename($filename) {
    $filename = preg_replace('/[\/\\\:*?"<>|]/', '-', $filename);
    $filename = preg_replace('/-+/', '-', $filename);
    if (strlen($filename) > 180) {
        $filename = substr($filename, 0, 180);
    }
    return $filename;
}

function ensureDirectoryExists($directory) {
    if (!is_dir($directory)) {
        if (!mkdir($directory, 0777, true)) {
            throw new Exception("Không thể tạo thư mục: $directory");
        }
    }
    return $directory;
}

function extractTextFromCell($cell) {
    $text = '';
    
    try {
        foreach ($cell->getElements() as $element) {
            if ($element instanceof \PhpOffice\PhpWord\Element\TextRun) {
                $text .= extractTextFromTextRun($element);
            } 
            elseif (method_exists($element, 'getText')) {
                $text .= $element->getText();
            }
            elseif (method_exists($element, 'getContent')) {
                $text .= $element->getContent();
            }
        }
    } catch (Exception $e) {
        return $text . '';
    }
    
    return trim($text);
}

function extractTextFromTextRun($textRun) {
    $text = '';
    
    try {
        foreach ($textRun->getElements() as $element) {
            if (method_exists($element, 'getText')) {
                $text .= $element->getText();
            } 
            elseif (method_exists($element, 'getContent')) {
                $text .= $element->getContent();
            }
            elseif ($element instanceof \PhpOffice\PhpWord\Element\TextRun) {
                $text .= extractTextFromTextRun($element);
            }
        }
    } catch (Exception $e) {
        return $text;
    }
    
    return $text;
}

function getTableSampleId($allSampleIds, $tableIndex, $tablesCount) {
    if (empty($allSampleIds)) {
        return 'UNKNOWN_' . time();
    }
    
    if (count($allSampleIds) == 1) {
        return $allSampleIds[0];
    }
    
    if (count($allSampleIds) == $tablesCount) {
        return $allSampleIds[$tableIndex];
    }
    
    $idIndex = min(floor($tableIndex * count($allSampleIds) / $tablesCount), count($allSampleIds) - 1);
    return $allSampleIds[$idIndex];
}

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    if (!isset($_FILES['word_files']) || empty($_FILES['word_files']['name'][0])) {
        die("Lỗi: Không có file nào được tải lên.");
    }

    $uploadDir = __DIR__ . '/output/';
    ensureDirectoryExists($uploadDir);

    $results = [];
    $totalTablesProcessed = 0;
    $debugInfo = '';
    
    for ($i = 0; $i < count($_FILES['word_files']['name']); $i++) {
        if ($_FILES['word_files']['error'][$i] !== UPLOAD_ERR_OK) {
            $results[] = [
                'filename' => $_FILES['word_files']['name'][$i],
                'status' => 'error',
                'message' => "Lỗi khi tải file lên. Mã lỗi: " . $_FILES['word_files']['error'][$i]
            ];
            continue;
        }

        $originalName = sanitizeFilename(pathinfo($_FILES['word_files']['name'][$i], PATHINFO_FILENAME));
        $uploadedFile = $uploadDir . basename($_FILES['word_files']['name'][$i]);

        if (!move_uploaded_file($_FILES['word_files']['tmp_name'][$i], $uploadedFile)) {
            $results[] = [
                'filename' => $_FILES['word_files']['name'][$i],
                'status' => 'error',
                'message' => "Không thể di chuyển file đã upload."
            ];
            continue;
        }

        try {
            // Bắt đầu transaction
            $pdo->beginTransaction();
            
            $phpWord = IOFactory::load($uploadedFile);
            $allSampleIds = extractAllSampleIds($phpWord);
            
            $debugInfo .= "<h3>Xử lý file: " . htmlspecialchars($_FILES['word_files']['name'][$i]) . "</h3>";
            $debugInfo .= "Mã mẫu tìm thấy: " . implode(', ', $allSampleIds) . "<br>";
            
            // Lưu thông tin file vào database
            $fileStmt = $pdo->prepare("INSERT INTO files (original_filename) VALUES (?)");
            $fileStmt->execute([$_FILES['word_files']['name'][$i]]);
            $fileId = $pdo->lastInsertId();
            
            // Lưu thông tin mẫu
            foreach ($allSampleIds as $sampleId) {
                $sampleStmt = $pdo->prepare("INSERT IGNORE INTO samples (sample_id, file_id) VALUES (?, ?)");
                $sampleStmt->execute([$sampleId, $fileId]);
            }

            $spreadsheet = new Spreadsheet();
            $allowedTitles = [
                'STT', 'TT', 'thông số phân tích', 'kết quả', 
                'chỉ tiêu', 'chỉ số', 'parameter', 'result', 
                'thông số', 'phép thử', 'test', 'đơn vị', 'unit'
            ];
            
            $spreadsheet->removeSheetByIndex(0);
            $validTables = [];
            
            // Tìm các bảng hợp lệ
            foreach ($phpWord->getSections() as $section) {
                foreach ($section->getElements() as $element) {
                    if (method_exists($element, 'getRows')) {
                        $rows = $element->getRows();
                        
                        if (count($rows) > 0) {
                            $firstRow = $rows[0];
                            $cells = $firstRow->getCells();
                            
                            if (count($cells) > 0) {
                                $firstCell = $cells[0];
                                $cellText = extractTextFromCell($firstCell);
                                
                                foreach ($allowedTitles as $title) {
                                    if (stripos($cellText, $title) !== false) {
                                        $validTables[] = $element;
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            
            $debugInfo .= "Số bảng hợp lệ tìm thấy: " . count($validTables) . "<br>";
            
            // Xử lý từng bảng
            foreach ($validTables as $tableIndex => $table) {
                $sampleId = getTableSampleId($allSampleIds, $tableIndex, count($validTables));
                $debugInfo .= "<h4>Bảng " . ($tableIndex + 1) . " - Mã mẫu: $sampleId</h4>";
                
                // Tạo sheet Excel
                $newSheet = new \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet($spreadsheet, "Bảng " . ($tableIndex + 1));
                $spreadsheet->addSheet($newSheet);
                $sheet = $spreadsheet->setActiveSheetIndex($tableIndex);
                
                $rowIndex = 1;
                $dataRowCounter = 0;
                $isFirstRow = true;
                $columnMap = [];
                
                foreach ($table->getRows() as $row) {
                    if ($isFirstRow) {
                        // Xử lý hàng tiêu đề
                        $cellCount = count($row->getCells());
                        $firstCellText = extractTextFromCell($row->getCells()[0]);
                        $sheet->setCellValueByColumnAndRow(1, $rowIndex, $firstCellText);
                        $sheet->setCellValueByColumnAndRow(2, $rowIndex, "Mã mẫu");
                        
                        // Tạo ánh xạ cột
                        for ($cellIndex = 1; $cellIndex < $cellCount; $cellIndex++) {
                            $cellText = extractTextFromCell($row->getCells()[$cellIndex]);
                            $sheet->setCellValueByColumnAndRow($cellIndex + 2, $rowIndex, $cellText);
                            
                            $fieldName = mapColumnToField($cellText);
                            if ($fieldName) {
                                $columnMap[$cellIndex - 1] = $fieldName; // -1 vì bỏ qua cột STT
                            }
                        }
                        
                        $debugInfo .= "Ánh xạ cột: " . json_encode($columnMap, JSON_UNESCAPED_UNICODE) . "<br>";
                        $isFirstRow = false;
                    } 
                    else {
                        // Xử lý hàng dữ liệu
                        $dataRowCounter++;
                        
                        // Excel
                        $sheet->setCellValueByColumnAndRow(1, $rowIndex, $dataRowCounter);
                        $sheet->setCellValueByColumnAndRow(2, $rowIndex, $sampleId);
                        
                        // Thu thập dữ liệu cho database
                        $rowData = [];
                        $cellIndex = 0;
                        foreach ($row->getCells() as $cell) {
                            if ($cellIndex > 0) { // Bỏ qua cột STT
                                $cellText = extractTextFromCell($cell);
                                $sheet->setCellValueByColumnAndRow($cellIndex + 2, $rowIndex, $cellText);
                                $rowData[] = $cellText;
                            }
                            $cellIndex++;
                        }
                        
                        // Chuẩn bị dữ liệu cho database
                        $insertData = [
                            'sample_id' => $sampleId,
                            'parameter' => '',
                            'result' => '',
                            'unit' => '',
                            'method' => '',
                            'detection_limit' => '',
                            'standard' => ''
                        ];
                        
                        // Ánh xạ dữ liệu theo cột
                        for ($j = 0; $j < count($rowData); $j++) {
                            if (isset($columnMap[$j]) && !empty($rowData[$j])) {
                                $fieldName = $columnMap[$j];
                                $insertData[$fieldName] = $rowData[$j];
                            }
                        }
                        
                        // Nếu không có ánh xạ, sử dụng thứ tự mặc định
                        if (empty($columnMap) && count($rowData) >= 2) {
                            $insertData['parameter'] = $rowData[0] ?? '';
                            $insertData['result'] = $rowData[1] ?? '';
                            $insertData['unit'] = $rowData[2] ?? '';
                        }
                        
                        // Lưu vào database nếu có dữ liệu
                        if (!empty($insertData['parameter']) || !empty($insertData['result'])) {
                            try {
                                $sql = "INSERT INTO analysis_results (sample_id, parameter, result, unit, method, detection_limit, standard) 
                                        VALUES (?, ?, ?, ?, ?, ?, ?)";
                                
                                $stmt = $pdo->prepare($sql);
                                $success = $stmt->execute([
                                    $insertData['sample_id'],
                                    $insertData['parameter'],
                                    $insertData['result'],
                                    $insertData['unit'],
                                    $insertData['method'],
                                    $insertData['detection_limit'],
                                    $insertData['standard']
                                ]);
                                
                                if ($success) {
                                    $insertId = $pdo->lastInsertId();
                                    $debugInfo .= "- ✅ Lưu hàng $dataRowCounter: ID $insertId, Parameter: '{$insertData['parameter']}', Result: '{$insertData['result']}'<br>";
                                }
                                
                            } catch (PDOException $e) {
                                $debugInfo .= "- ❌ Lỗi lưu hàng $dataRowCounter: " . $e->getMessage() . "<br>";
                            }
                        }
                    }
                    $rowIndex++;
                }
            }
            
            // Tạo file Excel
            if (count($validTables) > 0) {
                $spreadsheet->setActiveSheetIndex(0);
                
                $timestamp = date('Y-m-d_H-i-s');
                $sanitizedIds = [];
                foreach ($allSampleIds as $id) {
                    $sanitizedIds[] = sanitizeFilename($id);
                }
                $sampleIdString = !empty($sanitizedIds) ? implode('-', $sanitizedIds) : 'no-id';
                
                $safeExcelFilename = sanitizeFilename($originalName) . '_' . sanitizeFilename($sampleIdString) . '_' . $timestamp . '.xlsx';
                $excelFile = $uploadDir . $safeExcelFilename;
                
                $writer = new Xlsx($spreadsheet);
                $writer->save($excelFile);
            }
            
            // Commit transaction
            $pdo->commit();
            
            // Kiểm tra dữ liệu đã lưu
            $debugInfo .= "<h4>Kiểm tra dữ liệu đã lưu:</h4>";
            foreach ($allSampleIds as $sampleId) {
                $countStmt = $pdo->prepare("SELECT COUNT(*) FROM analysis_results WHERE sample_id = ?");
                $countStmt->execute([$sampleId]);
                $recordCount = $countStmt->fetchColumn();
                $debugInfo .= "- Mẫu $sampleId: $recordCount bản ghi<br>";
            }

            $results[] = [
                'filename' => $_FILES['word_files']['name'][$i],
                'status' => 'success',
                'tables_count' => count($validTables),
                'sample_ids' => !empty($allSampleIds) ? implode(', ', $allSampleIds) : 'Không tìm thấy',
                'excel_file' => isset($safeExcelFilename) ? $safeExcelFilename : null,
                'debug_info' => $debugInfo
            ];
            
            $totalTablesProcessed += count($validTables);

        } catch (Exception $e) {
            $pdo->rollback();
            $results[] = [
                'filename' => $_FILES['word_files']['name'][$i],
                'status' => 'error',
                'message' => "Lỗi khi xử lý file: " . $e->getMessage()
            ];
        }
    }

    // Hiển thị kết quả
    echo "<h2>Kết quả chuyển đổi</h2>";
    echo "<p>Đã xử lý " . count($results) . " file với tổng cộng " . $totalTablesProcessed . " bảng được trích xuất.</p>";
    
    foreach ($results as $result) {
        if ($result['status'] === 'success') {
            echo "<div style='border: 1px solid #ddd; padding: 10px; margin: 10px 0;'>";
            echo "<h3>✅ " . htmlspecialchars($result['filename']) . "</h3>";
            echo "<p><strong>Mã mẫu:</strong> " . htmlspecialchars($result['sample_ids']) . "</p>";
            echo "<p><strong>Số bảng:</strong> " . $result['tables_count'] . "</p>";
            if ($result['excel_file']) {
                echo "<p><a href='output/" . $result['excel_file'] . "'>📥 Tải về Excel</a></p>";
            }
            echo "<details><summary>Chi tiết xử lý</summary>" . $result['debug_info'] . "</details>";
            echo "</div>";
        } else {
            echo "<div style='border: 1px solid #f00; padding: 10px; margin: 10px 0;'>";
            echo "<h3>❌ " . htmlspecialchars($result['filename']) . "</h3>";
            echo "<p>" . htmlspecialchars($result['message']) . "</p>";
            echo "</div>";
        }
    }
    
    // Hiển thị dữ liệu mới nhất từ database
    echo "<h3>10 bản ghi mới nhất trong database:</h3>";
    try {
        $recentStmt = $pdo->query("SELECT ar.*, s.file_id FROM analysis_results ar 
                                   LEFT JOIN samples s ON ar.sample_id = s.sample_id 
                                   ORDER BY ar.id DESC LIMIT 10");
        echo "<table border='1' style='border-collapse: collapse; width: 100%;'>";
        echo "<tr><th>ID</th><th>Mã mẫu</th><th>Thông số</th><th>Kết quả</th><th>Đơn vị</th></tr>";
        while ($row = $recentStmt->fetch(PDO::FETCH_ASSOC)) {
            echo "<tr>";
            echo "<td>" . $row['id'] . "</td>";
            echo "<td>" . htmlspecialchars($row['sample_id']) . "</td>";
            echo "<td>" . htmlspecialchars($row['parameter']) . "</td>";
            echo "<td>" . htmlspecialchars($row['result']) . "</td>";
            echo "<td>" . htmlspecialchars($row['unit']) . "</td>";
            echo "</tr>";
        }
        echo "</table>";
    } catch (Exception $e) {
        echo "Lỗi hiển thị dữ liệu: " . $e->getMessage();
    }
    
    echo "<p><a href='index.php'>← Quay lại trang chủ</a></p>";
} else {
    echo "Truy cập không hợp lệ.";
}

if (isset($tempDir)) {
    cleanupTempFiles($tempDir);
}
?>