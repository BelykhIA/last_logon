# XML запит для фільтрації подій
$query = @'
<QueryList>
<Query Id='0' Path='Security'>
<Select Path='Security'>
*[System[EventID='4624']
and(
EventData[Data[@Name='VirtualAccount']='%%1843']
and
(EventData[Data[@Name='LogonType']='2'] or EventData[Data[@Name='LogonType']='7'])
)
]
</Select>
</Query>
</QueryList>
'@

# Отримання подій
$events = Get-WinEvent -FilterXml $query

# Отримання імені комп'ютера
$computerName = $env:COMPUTERNAME

# Обробка даних подій з різними властивостями для LogonType 2 і 7
$processedEvents = $events | ForEach-Object {
    $event = $_
    $logonType = $event.Properties[8].Value

    if ($logonType -eq '2') {
        [PSCustomObject]@{
            User      = $event.Properties[1].Value # Ім'я користувача для LogonType 2
            Domain    = $event.Properties[2].Value # Домен для LogonType 2
            TimeStamp = $event.TimeCreated # Час створення події
            LogonType = $logonType # Тип входу
            Computer  = $computerName # Ім'я комп'ютера
        }
    } elseif ($logonType -eq '7') {
        [PSCustomObject]@{
            User      = $event.Properties[5].Value # Ім'я користувача для LogonType 7
            Domain    = $event.Properties[6].Value # Домен для LogonType 7
            TimeStamp = $event.TimeCreated # Час створення події
            LogonType = $logonType # Тип входу
            Computer  = $computerName # Ім'я комп'ютера
        }
    }
}

# Шлях до тимчасового CSV файлу
$tempCsvPath = "D:\Inventory\Login\$computerName.csv"
$processedEvents | Export-Csv -Path $tempCsvPath -NoTypeInformation -Encoding UTF8

# Видалення дублікованих записів
$uniqueEvents = Import-Csv -Path $tempCsvPath | Sort-Object -Property TimeStamp, User -Unique
$uniqueCsvPath = "D:\Inventory\Login\$computerName.csv"
$uniqueEvents | Export-Csv -Path $uniqueCsvPath -NoTypeInformation -Encoding UTF8

# Формування шляху для збереження Excel файлу
$excelPath = "D:\Inventory\Login\LoginPC.xlsx"

# Конвертація CSV у формат Excel
try {
    # Запуск Excel
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false

    # Відкриття або створення робочої книги
    if (Test-Path $excelPath) {
        $workbook = $excel.Workbooks.Open($excelPath)
    } else {
        $workbook = $excel.Workbooks.Add()
    }

    # Перевірка наявності аркуша з ім'ям комп'ютера
    $sheetExists = $false
    foreach ($sheet in $workbook.Sheets) {
        if ($sheet.Name -eq $computerName) {
            $sheetExists = $true
            $worksheet = $sheet
            break
        }
    }

    # Якщо аркуша не існує, створити новий
    if (-not $sheetExists) {
        $worksheet = $workbook.Sheets.Add()
        $worksheet.Name = $computerName
    }

    # Завантаження даних з CSV
    $csvContent = Import-Csv -Path $uniqueCsvPath

    # Знаходження останнього рядка з даними
    $lastRow = $worksheet.UsedRange.Rows.Count

    # Додавання заголовків, якщо аркуш новий
    if ($lastRow -eq 1 -and $worksheet.Cells.Item(1, 1).Value2 -eq $null) {
        $headers = $csvContent[0].PSObject.Properties.Name
        $col = 1
        foreach ($header in $headers) {
            $worksheet.Cells.Item(1, $col).Value2 = $header
            $col++
        }
        $lastRow++
    }

    # Додавання даних
    $row = $lastRow + 1
    foreach ($entry in $csvContent) {
        $col = 1
        foreach ($header in $headers) {
            $worksheet.Cells.Item($row, $col).Value2 = $entry.$header
            $col++
        }
        $row++
    }

    # Збереження файлу Excel
    $excel.DisplayAlerts = $false # Вимкнення попереджень
    $workbook.SaveAs($excelPath, 51) # 51 = Excel XLSX format
    $workbook.Close()
    $excel.Quit()
} catch {
    Write-Output "Error: $_" > "D:\Inventory\Login\$computerName.txt"
} finally {
    # Очистка COM об'єктів
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    # Видалення тимчасових файлів
    Remove-Item -Path $tempCsvPath -Force
    if (Test-Path $uniqueCsvPath) {
        Remove-Item -Path $uniqueCsvPath -Force
    }
}
