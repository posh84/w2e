<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <title>Upload Word và Xuất Excel</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            max-width: 800px;
            margin: 0 auto;
            padding: 20px;
        }
        h2 {
            color: #333;
        }
        .form-group {
            margin-bottom: 15px;
        }
        button {
            background-color: #4CAF50;
            color: white;
            padding: 10px 15px;
            border: none;
            border-radius: 4px;
            cursor: pointer;
        }
        button:hover {
            background-color: #45a049;
        }
    </style>
</head>
<body>
    <h2>Upload file Word (.doc hoặc .docx)</h2>
    <form method="POST" action="upload.php" enctype="multipart/form-data">
        <div class="form-group">
            <input type="file" name="word_files[]" accept=".doc,.docx" multiple required>
            <small>Có thể chọn nhiều file cùng lúc</small>
        </div>
        <button type="submit">Chuyển đổi</button>
    </form>
</body>
</html>