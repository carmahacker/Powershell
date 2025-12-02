# Email configuration
$emailSubjectPrefix = "Отчёт о резервном копировании"

# Параметры подключения
$veeamServers = @(
    @{ Address = "veeam3-spb-79.pharmasyntez.com"; Description = "Фармасинтез-Норд, 36 участок" },
    @{ Address = "spb-veeam.pharmasyntez.com"; Description = "Фармасинтез-Норд, ЦТТ" }
)

# Описание задач и адреса электронной почты
$jobDescriptions = @(
    @{ Name = "Lab-FHL1"; Description = "Фармасинтез-Норд, ЦТТ"; Email = @("spb.veeam.fhl1@pharmasyntez.com") },
    @{ Name = "Lab-BC-FHL"; Description = "Фармасинтез-Биокапитал"; Email = @("bc.veeam.fhl@pharmasyntez.com") },
    @{ Name = "Lab-FHL2"; Description = "Фармасинтез-Норд, 36 участок"; Email = @("spb.veeam.fhl2@pharmasyntez.com") },
    @{ Name = "LAB-BAL"; Description = "Фармасинтез-Норд, 36 участок"; Email = @("spb.veeam.bal@pharmasyntez.com") }
)

$username = "veeam_operator"
$password = "1qa11111"

# Function to send email
function funGAM_SendEmail {
    param (
        [string[]]$inEmailTo,
        [string]$inEmailSubject,
        [string]$inEmailBody
    )

    [string]$EmailServer = "ex1.pharmasyntez.com"
    [string]$EmailFrom = "helpdesk_service@pharmasyntez.com"
    $EmailPort = 2255

    $smtpClient = New-Object System.Net.Mail.SmtpClient($EmailServer, $EmailPort)
    $smtpClient.EnableSsl = $false
    
    $mailMessage = New-Object System.Net.Mail.MailMessage
    $mailMessage.From = $EmailFrom
    $mailMessage.Subject = $inEmailSubject
    $mailMessage.Body = $inEmailBody
    $mailMessage.IsBodyHtml = $true

    foreach ($email in $inEmailTo) {
        $mailMessage.To.Add($email)
    }

    try {
        $smtpClient.Send($mailMessage)
        Write-Host "Письмо успешно отправлено на $($inEmailTo -join ', ')" -ForegroundColor Green
    } catch {
        Write-Host "Ошибка при отправке письма на $($inEmailTo -join ', '): $_" -ForegroundColor Red
    } finally {
        $mailMessage.Dispose()
        $smtpClient.Dispose()
    }
}

# Проверка и загрузка модуля Veeam.Backup.PowerShell
$veeamModule = "Veeam.Backup.PowerShell"
try {
    if (-not (Get-Module -ListAvailable -Name $veeamModule)) {
        Write-Error "Модуль $veeamModule не найден. Убедитесь, что Veeam Backup & Replication PowerShell модуль установлен."
        exit
    }
    
    if (-not (Get-Module -Name $veeamModule)) {
        Import-Module -Name $veeamModule -ErrorAction Stop
        Write-Host "Модуль $veeamModule успешно загружен."
    }
}
catch {
    Write-Error "Ошибка при загрузке модуля $veeamModule $($_.Exception.Message)"
    exit
}

# Создаем учетные данные (общие для всех серверов)
$securePassword = ConvertTo-SecureString $password -AsPlainText -Force
$credential = New-Object System.Management.Automation.PSCredential($username, $securePassword)

try {
    foreach ($veeamServer in $veeamServers) {
        $serverAddress = $veeamServer.Address
        $serverDescription = $veeamServer.Description
        Write-Host "Обработка сервера: $serverAddress ($serverDescription)"

        try {
            # Подключаемся к Veeam серверу
            Connect-VBRServer -Server $serverAddress -Credential $credential -ErrorAction Stop
            Write-Host "Успешно подключено к Veeam серверу: $serverAddress"
        }
        catch {
            Write-Error "Ошибка подключения к Veeam серверу $serverAddress : $($_.Exception.Message)"
            continue
        }

        try {
            # Получаем все задачи
            $jobs = Get-VBRComputerBackupJob | Where-Object { $_.Name -like "LAB-*" }
            if (-not $jobs) {
                Write-Warning "Задачи не найдены через Get-VBRComputerBackupJob на сервере $serverAddress. Пробуем Get-VBRJob..."
                $jobs = Get-VBRJob | Where-Object { $_.Name -like "LAB-*" }
                if (-not $jobs) {
                    Write-Error "Задачи с именами, начинающимися на LAB-, не найдены на сервере $serverAddress!"
                    continue
                }
            }

            # Выводим в консоль найденные имена задач
            Write-Host "Найденные задачи, начинающиеся с 'LAB-' на сервере $serverAddress :"
            $jobs | ForEach-Object { Write-Host $_.Name -ForegroundColor Cyan }

            # Обработка каждой задачи
            foreach ($job in $jobs) {
                $jobName = $job.Name
                # Получаем описание и email задачи
                $jobInfo = $jobDescriptions | Where-Object { $_.Name -eq $jobName }
                $jobDescription = $jobInfo.Description
                $emailRecipients = $jobInfo.Email
                if (-not $jobDescription) {
                    $jobDescription = "Описание не указано"
                }
                if (-not $emailRecipients) {
                    $emailRecipients = @("it.spb@pharmasyntez.com")
                    Write-Warning "Email для задачи $jobName не указан, используется адрес по умолчанию."
                }
                Write-Host "Обработка задачи: $jobName ($jobDescription) на сервере $serverAddress"

                # Получаем последнюю сессию для задачи
                $lastSession = $job.FindLastSession()
                if (-not $lastSession) {
                    Write-Warning "Последняя сессия для задачи $jobName не найдена на сервере $serverAddress."
                    continue
                }

                $result = @()

                # Получаем задачи из сессии
                $taskSessions = Get-VBRTaskSession -Session $lastSession
                if (-not $taskSessions) {
                    Write-Warning "Задачи в сессии от $($lastSession.CreationTime) не найдены на сервере $serverAddress."
                    continue
                }

                foreach ($task in $taskSessions) {
                    $serverName = $task.Name
                    Write-Host "Обработка сервера: $serverName"

                    # Получаем описание ManagedServer
                    $managedServer = Get-VBRServer -Name $serverName -ErrorAction SilentlyContinue
                    $description = if ($managedServer) { $managedServer.Description } else { "Not found in Veeam infrastructure" }

                    # Изменяем статус Warning и Success на OK
                    $status = if ($task.Status -eq "Warning" -or $task.Status -eq "Success") { "OK" } else { $task.Status }

                    $result += [PSCustomObject]@{
                        SessionDate       = $lastSession.CreationTime.ToString("dd MMMM yyyy г. HH:mm:ss")
                        ServerName        = $serverName
                        Description       = $description
                        Status            = $status
                        StartTime         = $task.Progress.StartTimeLocal.ToString("HH:mm:ss")
                        EndTime           = $task.Progress.StopTimeLocal.ToString("HH:mm:ss")
                        Duration          = ($task.Progress.StopTimeLocal - $task.Progress.StartTimeLocal).ToString("hh\:mm\:ss")
                        BackedUpFiles     = $task.Progress.ProcessedObjects
                        Size              = if ($task.Progress.ProcessedSize -ge 1GB) {
                            "{0:N1} GB" -f ($task.Progress.ProcessedSize / 1GB)
                        } elseif ($task.Progress.ProcessedSize -ge 1MB) {
                            "{0:N1} MB" -f ($task.Progress.ProcessedSize / 1MB)
                        } else {
                            "{0:N1} B" -f $task.Progress.ProcessedSize
                        }
                    }
                }

                # Вывод в консоль и отправка по email
                if ($result) {
                    Write-Output "File Backup job: $jobName ($jobDescription) on server $serverAddress"
                    foreach ($sessionGroup in ($result | Group-Object SessionDate)) {
                        Write-Output ($sessionGroup.Group | Select-Object -First 1).SessionDate
                        Write-Output "Details"
                        Write-Output "Name`tDescription`tStatus`tStart time`tEnd time`tDuration`tBacked up files`tSize"
                        foreach ($item in $sessionGroup.Group) {
                            Write-Output "$($item.ServerName)`t$($item.Description)`t$($item.Status)`t$($item.StartTime)`t$($item.EndTime)`t$($item.Duration)`t$($item.BackedUpFiles)`t$($item.Size)"
                        }
                        Write-Output ""
                    }

                    # Сохранение в CSV
                    $csvFileName = "VeeamReport_${jobName}_${serverAddress}_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
                    $result | Select-Object ServerName, Description, Status, StartTime, EndTime, Duration, BackedUpFiles, Size | Export-Csv -Path $csvFileName -NoTypeInformation -Encoding UTF8
                    Write-Host "Данные для задачи $jobName на сервере $serverAddress сохранены в $csvFileName"

                    # Формирование HTML для email
                    $htmlBody = @"
<html>
<head>
    <style>
        body { font-family: Arial, sans-serif; }
        h2 { color: #2e6c80; }
        table { border-collapse: collapse; width: 100%; margin-top: 10px; }
        th, td { border: 1px solid #dddddd; text-align: left; padding: 8px; }
        th { background-color: #f2f2f2; }
        tr:nth-child(even) { background-color: #f9f9f9; }
    </style>
</head>
<body>
    <h2>Отчёт о резервном копировании: $jobName ($jobDescription)</h2>
    <p><strong>Дата:</strong> $(($result | Select-Object -First 1).SessionDate)</p>
    <table>
        <tr>
            <th>Имя ПК</th>
            <th>Прибор</th>
            <th>Статус</th>
            <th>Начало</th>
            <th>Завершение</th>
            <th>Продолжительность</th>
            <th>Размер</th>
        </tr>
"@

                    foreach ($item in $result) {
                        $htmlBody += @"
        <tr>
            <td>$($item.ServerName)</td>
            <td>$($item.Description)</td>
            <td>$($item.Status)</td>
            <td>$($item.StartTime)</td>
            <td>$($item.EndTime)</td>
            <td>$($item.Duration)</td>
            <td>$($item.Size)</td>
        </tr>
"@
                    }

                    $htmlBody += @"
    </table>
</body>
</html>
"@

                    # Отправка email
                    $emailSubject = "$emailSubjectPrefix - $jobName ($jobDescription) - $(Get-Date -Format 'yyyy-MM-dd')"
                    funGAM_SendEmail -inEmailTo $emailRecipients -inEmailSubject $emailSubject -inEmailBody $htmlBody
                }
                else {
                    Write-Warning "Нет данных для задачи $jobName на сервере $serverAddress."
                }
            }
        }
        catch {
            Write-Error "Ошибка при выполнении скрипта на сервере $serverAddress : $($_.Exception.Message)"
        }
        finally {
            # Отключаемся от Veeam сервера
            Disconnect-VBRServer
            Write-Host "Отключено от Veeam сервера $serverAddress."
        }
    }
}
catch {
    Write-Error "Общая ошибка при выполнении скрипта: $($_.Exception.Message)"
}
