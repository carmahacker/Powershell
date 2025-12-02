# =====================================================================
#     МАССИВ СХД — теперь можно указывать любые IP
# =====================================================================
$storage = @(
    @{
        IP       = "10.254.250.105"
        Login    = "administrator"
        Password = '444444444
    },
    @{
        IP       = "10.254.250.180"
        Login    = "manager"
        Password = '444444444'
    }
)
# =====================================================================
#     СПРАВОЧНИК ГОРОДОВ
# =====================================================================
function Get-CityByIP {
    param([string]$ip)

    if ($ip -match '^10\.25\.') { return 'Уссурийск' }
    if ($ip -match '^10\.254\.') { return 'Москва-ЦОД' }
    if ($ip -match '^10\.79\.')  { return 'Санкт-Петербург 36' }
    if ($ip -match '^10\.78\.')  { return 'Санкт-Петербург ЦТТ' }
    if ($ip -match '^10\.38\.')  { return 'Братск' }
    if ($ip -match '^10\.72\.')  { return 'Тюмень' }
    if ($ip -match '^10\.1\.')   { return 'Иркутск' }
    if ($ip -match '^10\.123\.')   { return 'Москва-Биокапитал' }
    return 'Неизвестно'
}

# =====================================================================
#        Функция: Перевод часов в годы
# =====================================================================
function Convert-HoursToYears {
    param([int]$hours)
    if ($hours -le 0) { return 0 }
    return [math]::Round($hours / 24 / 365, 2)
}

# =====================================================================
#        Функция: Парсер XML-like PROPERTY-тегов
# =====================================================================
function Parse-PropertyXml {
    param([string[]]$lines)

    $results = @()
    $current = [ordered]@{}

    foreach ($line in $lines) {

        if ($line -match '<OBJECT[^>]+basetype="drives"') {
            if ($current.Count -gt 0) {
                $results += [pscustomobject]$current
                $current = [ordered]@{}
            }
            continue
        }

        if ($line -match '<PROPERTY name="([^"]+)"[^>]*>(.*?)</PROPERTY>') {
            $name  = $Matches[1]
            $value = $Matches[2].Trim()

            switch ($name) {
                "location"       { $current.Location = $value }
                "serial-number"  { $current.Serial   = $value }
                "vendor"         { $current.Vendor   = $value }
                "size"           { $current.Size     = $value }
                "health"         { $current.Health   = $value }
                "power-on-hours" {
                    $current.PowerOnHours = $value
                    $current.PowerOnYears = Convert-HoursToYears $value
                }
            }
        }
    }

    if ($current.Count -gt 0) {
        $results += [pscustomobject]$current
    }

    return $results
}

# =====================================================================
#        Сбор данных
# =====================================================================
$final = @()

foreach ($st in $storage) {

    Write-Host "Подключение к $($st.IP)..." -ForegroundColor Cyan

    $cred = New-Object Management.Automation.PSCredential(
        $st.Login,
        (ConvertTo-SecureString $st.Password -AsPlainText -Force)
    )

    try {
        $ssh = New-SSHSession -ComputerName $st.IP -Credential $cred -AcceptKey -ErrorAction Stop
    }
    catch {
        Write-Warning "Не удалось подключиться к $($st.IP): $_"
        continue
    }

    try {
        $cmd = Invoke-SSHCommand -SSHSession $ssh -Command "show disks detail" -ErrorAction Stop
    }
    catch {
        Write-Warning "Ошибка выполнения команды на $($st.IP): $_"
        Remove-SSHSession -SSHSession $ssh
        continue
    }

    Remove-SSHSession -SSHSession $ssh

    $parsed = Parse-PropertyXml -lines $cmd.Output

    foreach ($row in $parsed) {

# Гарантируем наличие всех нужных свойств и корректные значения по умолчанию
$defaultProps = @{
    Location      = ""
    Serial        = ""
    Vendor        = ""
    Size          = ""
    Health        = ""
    PowerOnHours  = 0
    PowerOnYears  = 0.0   # важно именно 0.0, а не 0, чтобы тип был double
}

foreach ($prop in $defaultProps.Keys) {
    if (-not $row.PSObject.Properties[$prop]) {
        $row | Add-Member -NotePropertyName $prop -NotePropertyValue $defaultProps[$prop] -Force
    }
    # Если свойство есть, но там null или пусто — ставим значение по умолчанию
    if ($null -eq $row.$prop -or $row.$prop -eq "") {
        $row.$prop = $defaultProps[$prop]
    }
}

        $enc = ""
        $slot = ""
        if ($row.Location -match "(\d+)\.(\d+)") {
            $enc  = $Matches[1]
            $slot = $Matches[2]
        }

        $city = Get-CityByIP $st.IP

        $final += [pscustomobject]@{
            StorageIP    = $st.IP
            City         = $city
            Enclosure    = $enc
            Slot         = $slot
            Location     = $row.Location
            Serial       = $row.Serial
            Vendor       = $row.Vendor
            Size         = $row.Size
            Health       = $row.Health
            PowerOnHours = $row.PowerOnHours
            PowerOnYears = $row.PowerOnYears
        }
    }
}

# сортировка
$final = $final | Sort-Object City, StorageIP, Enclosure, Slot


# =====================================================================
#        Нормализация моделей
# =====================================================================
$normalized = foreach ($disk in $final) {

    $row = [pscustomobject]@{
        StorageIP    = $disk.StorageIP
        City         = $disk.City
        Enclosure    = $disk.Enclosure
        Slot         = $disk.Slot
        Location     = $disk.Location
        Serial       = $disk.Serial
        Vendor       = $disk.Vendor
        Size         = $disk.Size
        Health       = $disk.Health
        PowerOnHours = $disk.PowerOnHours
        PowerOnYears = $disk.PowerOnYears
        SizeGB       = 0
        ModelNorm    = ""
    }

    if ($row.Size -match "([0-9\.]+)GB") {
        $row.SizeGB = [double]$Matches[1]
    }

    if (
        ($row.Vendor -eq "SEAGATE" -and $row.SizeGB -lt 1500) -or
        ($row.Vendor -eq "SAMSUNG")
    ) {
        $row.Vendor = "TOSHIBA"
        $row.SizeGB = 2400.4
        $row.Size   = "2400.4GB"
    }

    $row.ModelNorm = "$($row.Vendor) $($row.SizeGB)GB"

    $row
}


# =====================================================================
#        Таблица Spare Requirements
# =====================================================================
$spare = $normalized |
    Group-Object ModelNorm, Enclosure |
    ForEach-Object {
        $sample = $_.Group[0]
        [pscustomobject]@{
            ModelNorm      = $sample.ModelNorm
            SizeGB         = $sample.SizeGB
            Enclosure      = $sample.Enclosure
            City           = $sample.City
            RequiredSpare  = 2
        }
    } | Sort-Object City, Enclosure, ModelNorm


# =====================================================================
#        Excel
# =====================================================================
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$wb = $excel.Workbooks.Add()

while ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item(1).Delete() }

# ---------------- Лист 1
$ws1 = $wb.Worksheets.Item(1)
$ws1.Name = "Full_Disks_List"

$headers = @(
    "City","StorageIP","Enclosure","Slot",
    "Serial","Vendor","ModelNorm","SizeGB",
    "PowerOnYears","Health"
)

$row = 1
$col = 1
foreach ($h in $headers) {
    $ws1.Cells.Item($row,$col).Value2 = $h
    $ws1.Cells.Item($row,$col).Font.Bold = $true
    $col++
}

$row = 2
foreach ($d in $normalized) {

    $ws1.Cells.Item($row,1).Value2 = $d.City
    $ws1.Cells.Item($row,2).Value2 = $d.StorageIP
    $ws1.Cells.Item($row,3).Value2 = $d.Enclosure
    $ws1.Cells.Item($row,4).Value2 = $d.Slot
    $ws1.Cells.Item($row,5).Value2 = $d.Serial
    $ws1.Cells.Item($row,6).Value2 = $d.Vendor
    $ws1.Cells.Item($row,7).Value2 = $d.ModelNorm
    $ws1.Cells.Item($row,8).Value2 = [double]$d.SizeGB
    $ws1.Cells.Item($row,9).Value2 = [double]$d.PowerOnYears
    $ws1.Cells.Item($row,10).Value2 = $d.Health

    # Подсветка > 3 лет
    if ($d.PowerOnYears -gt 3) {
        $ws1.Range("A$row","J$row").Font.Bold = $true
        $ws1.Range("A$row","J$row").Font.ColorIndex = 3
    }

    # Подсветка Fault
    if ($d.Health -eq "Fault") {
        $ws1.Range("A$row","J$row").Interior.ColorIndex = 7   # фиолетовый
        $ws1.Range("A$row","J$row").Font.Bold = $true
    }

    $row++
}

$ws1.Columns.AutoFit()

# ---------------- Лист 2
$ws2 = $wb.Worksheets.Add()
$ws2.Name = "Spare_Requirements"

$headers2 = @("City","ModelNorm","SizeGB","Enclosure","RequiredSpare")
$row = 1
$col = 1
foreach ($h in $headers2) {
    $ws2.Cells.Item($row,$col).Value2 = $h
    $ws2.Cells.Item($row,$col).Font.Bold = $true
    $col++
}

$row = 2
foreach ($s in $spare) {
    $ws2.Cells.Item($row,1).Value2 = $s.City                # строка — ок
    $ws2.Cells.Item($row,2).Value2 = $s.ModelNorm           # строка — ок
    $ws2.Cells.Item($row,3).Value2 = [decimal]$s.SizeGB     # ← исправлено
    $ws2.Cells.Item($row,4).Value2 = $s.Enclosure           # строка или число — ок
    $ws2.Cells.Item($row,5).Value2 = [decimal]2             # ← тоже через decimal (или просто 2 — обычно проходит)

    $row++
}
$ws2.Columns.AutoFit()

# Сохранение
$folder = "C:\Reports"
if (-not (Test-Path $folder)) {
    New-Item -Path $folder -ItemType Directory | Out-Null
}
$path = "C:\Reports\Storage_Report_{0}.xlsx" -f (Get-Date -Format "yyyyMMdd_HHmmss")
$wb.SaveAs($path)
$excel.Quit()

Write-Host "`nГотово! Отчёт сформирован:" -ForegroundColor Green
Write-Host $path
