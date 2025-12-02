# Устанавливаем кодировку 
[Console]::InputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

# Конфигурация
$botToken = "7917974357:AAEA4vZ_Sxa0-28F1111111111111111"  # Укажите токен вашего бота
$apiBaseUrl = "https://api.telegram.org/bot$botToken"
$logDir = "C:\infrabot\locker_logs"  # Абсолютный путь для папки логов
$maxRetries = 5  # Максимальное количество попыток при сбоях
$baseRetryDelay = 1  # Базовая задержка в секундах перед повторной попыткой
$timeoutSeconds = 2  # Таймаут для запросов к Telegram API

# Список разрешённых UserId
$allowedUserIds = @(
    287211111,  # Печковский
    172911111
)

# Функция для получения пути к текущему лог-файлу
function Get-CurrentLogFile {
    $currentDate = (Get-Date).ToString("yyyy-MM-dd")
    $logFileName = "bot_log_$currentDate.txt"
    return Join-Path -Path $logDir -ChildPath $logFileName
}

# Функция для проверки и создания папки логов
function Ensure-LogDirectory {
    try {
        # Проверка, существует ли папка
        if (-not (Test-Path -Path $logDir -PathType Container)) {
            New-Item -Path $logDir -ItemType Directory -Force -ErrorAction Stop | Out-Null
        }
        # Проверка записи в папку
        $testFile = Join-Path -Path $logDir -ChildPath "test.txt"
        [System.IO.File]::WriteAllText($testFile, "test", [System.Text.Encoding]::UTF8) | Out-Null
        Remove-Item -Path $testFile -Force -ErrorAction SilentlyContinue
        return $true
    } catch {
        Write-Host "[ERROR] $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') Не удалось создать или проверить папку $logDir : $_" -ForegroundColor Red
        return $false
    }
}

# Функция для проверки доступности лог-файла
function Test-LogFile {
    $logFile = Get-CurrentLogFile
    try {
        if (-not (Test-Path -Path $logFile)) {
            New-Item -Path $logFile -ItemType File -Force -ErrorAction Stop | Out-Null
        }
        # Проверка записи в лог-файл
        [System.IO.File]::AppendAllText($logFile, "", [System.Text.Encoding]::UTF8) | Out-Null
        return $true
    } catch {
        Write-Host "[ERROR] $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') Не удалось создать или записать в лог-файл $logFile : $_" -ForegroundColor Red
        return $false
    }
}

# Функция для записи в лог
function Write-Log {
    param (
        [string]$Message,
        [switch]$IsError
    )
    if (-not (Ensure-LogDirectory) -or -not (Test-LogFile)) {
        if ($IsError) {
            Write-Host "[ERROR] $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') Лог-файл недоступен. Сообщение: $Message" -ForegroundColor Red
        }
        return
    }
    $logFile = Get-CurrentLogFile
    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $logEntry = if ($IsError) { "[ERROR] $timestamp $Message" } else { "[INFO] $timestamp $Message" }
    try {
        [System.IO.File]::AppendAllText($logFile, "$logEntry`n", [System.Text.Encoding]::UTF8) | Out-Null
        if ($IsError) {
            Write-Host $logEntry -ForegroundColor Red
        }
    } catch {
        Write-Host "[ERROR] $timestamp Не удалось записать в лог-файл $logFile : $_" -ForegroundColor Red
    }
}

# Функция для проверки доступности Telegram API
function Test-TelegramApi {
    try {
        $response = Invoke-WebRequest -Uri "$apiBaseUrl/getMe" -Method Get -TimeoutSec $timeoutSeconds -ErrorAction Stop
        return $response.StatusCode -eq 200
    } catch {
        # Игнорируем ошибку таймаута
        if ($_.Exception.Message -notlike "*The request was canceled due to the configured HttpClient.Timeout*") {
            Write-Log "Ошибка проверки Telegram API: $_" -IsError
        }
        return $false
    }
}

# Функция для получения displayName из Active Directory
function Get-ADUserDisplayName {
    param (
        [Parameter(Mandatory)] [string]$UserId
    )
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
        $adUser = Get-ADUser -Filter "extensionAttribute6 -eq '$UserId'" -Properties displayName, extensionAttribute6 -ErrorAction Stop
        if ($adUser) {
            return $adUser.displayName
        } else {
            Write-Log "Пользователь с extensionAttribute6=$UserId не найден в Active Directory." -IsError
            return "Неизвестный пользователь ($UserId)"
        }
    } catch {
        Write-Log "Ошибка при получении displayName для UserId $UserId : $_" -IsError
        return "Ошибка получения имени ($UserId)"
    }
}

# Функция для отправки сообщений в Telegram с повторными попытками
function Send-TelegramMessage {
    param (
        [Parameter(Mandatory)] [string]$ChatId,
        [Parameter(Mandatory)] [string]$Text,
        [hashtable]$ReplyMarkup = $null,
        [string]$ThreadId = $null
    )

    if ([string]::IsNullOrEmpty($ChatId)) {
        Write-Log "Ошибка: ChatId пустой при попытке отправки сообщения." -IsError
        return $null
    }

    $body = @{
        chat_id = $ChatId
        text    = $Text
    }

    # Добавляем message_thread_id для супергрупп
    if ($ChatId -eq "-1002173953951") {
        $body.message_thread_id = 112218
    } elseif ($ChatId -eq "-1002241626299") {
        $body.message_thread_id = 35683
    } elseif ($ThreadId) {
        $body.message_thread_id = $ThreadId
    }

    if ($ReplyMarkup) {
        $body.reply_markup = $ReplyMarkup | ConvertTo-Json -Depth 3 -Compress
    }

    $retryCount = 0
    while ($retryCount -lt $maxRetries) {
        try {
            $response = Invoke-RestMethod -Uri "$apiBaseUrl/sendMessage" -Method Post -ContentType "application/json" -Body ($body | ConvertTo-Json -Depth 3 -Compress) -TimeoutSec $timeoutSeconds -ErrorAction Stop
            Write-Log "Сообщение отправлено в чат $ChatId (thread_id: $($body.message_thread_id)) : $Text"
            return $response
        } catch {
            $retryCount++
            if ($retryCount -eq $maxRetries -and $_.Exception.Message -notlike "*The request was canceled due to the configured HttpClient.Timeout*") {
                Write-Log "Достигнуто максимальное количество попыток отправки сообщения в чат $ChatId : $_" -IsError
                return $null
            }
            $delay = $baseRetryDelay * [math]::Pow(2, $retryCount - 1)
            Start-Sleep -Seconds $delay
        }
    }
}

# Функция для получения заблокированных учетных записей
function Get-LockedUsers {
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
        $forest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
        $domains = $forest.Domains
        $lockedUsers = @()

        foreach ($domain in $domains) {
            $domainController = $domain.PdcRoleOwner.Name
            $lockedUsersInDomain = Search-ADAccount -LockedOut -UsersOnly -Server $domainController -ErrorAction Stop
            $lockedUsers += $lockedUsersInDomain
        }

        return $lockedUsers
    } catch {
        Write-Log "Ошибка при получении заблокированных учетных записей: $_" -IsError
        return @()
    }
}

# Функция для разблокировки учетной записи
function Unlock-UserAccount {
    param (
        [Parameter(Mandatory)] [string]$SamAccountName
    )

    try {
        Import-Module ActiveDirectory -ErrorAction Stop
        Unlock-ADAccount -Identity $SamAccountName -ErrorAction Stop
        Write-Log "Учетная запись $SamAccountName успешно разблокирована."
        return "Учетная запись $SamAccountName успешно разблокирована."
    } catch {
        Write-Log "Ошибка при разблокировке учетной записи $SamAccountName : $_" -IsError
        return "Ошибка при разблокировке учетной записи $SamAccountName."
    }
}

# Функция для создания кнопок с пользовательскими данными
function Create-TelegramButtons {
    param (
        [Parameter(Mandatory)] [array]$UserList
    )

    $inlineKeyboard = @()

    foreach ($user in $UserList) {
        $userName = $user.Name
        if ([string]::IsNullOrEmpty($userName)) {
            $userName = "Без имени"
        }

        $button = @{
            text = $userName
            callback_data = $user.SamAccountName
        }

        $inlineKeyboard += ,@($button)
    }

    return @{
        inline_keyboard = $inlineKeyboard
    }
}

# Функция для обработки обновлений
function Handle-Update {
    param (
        [Parameter(Mandatory)] [PSCustomObject]$Update
    )

    $userId = $null
    $chatId = $null
    $chatType = $null
    $text = $null
    $displayName = $null

    # Проверка корректности обновления
    if (-not $Update -or -not $Update.PSObject.Properties["update_id"]) {
        Write-Log "Получено некорректное обновление." -IsError
        return
    }

    # Определяем userId и chatId
    if ($Update.PSObject.Properties["message"] -and $Update.message) {
        $userId = $Update.message.from.id
        $chatId = $Update.message.chat.id
        $chatType = $Update.message.chat.type
        $text = $Update.message.text
    } elseif ($Update.PSObject.Properties["callback_query"] -and $Update.callback_query) {
        $userId = $Update.callback_query.from.id
        $chatId = $Update.callback_query.message.chat.id
        $chatType = $Update.callback_query.message.chat.type
    } elseif ($Update.PSObject.Properties["edited_message"]) {
        Write-Log "Получено отредактированное сообщение от userId $($Update.edited_message.from.id). Игнорируется." -IsError
        return
    } else {
        Write-Log "Неизвестный тип обновления от update_id $($Update.update_id)." -IsError
        return
    }

    # Проверка на наличие userId и chatId
    if ([string]::IsNullOrEmpty($userId) -or [string]::IsNullOrEmpty($chatId)) {
        Write-Log "Ошибка: userId ($userId) или chatId ($chatId) не определены." -IsError
        return
    }

    # Проверка разрешённого UserId
    if ($userId -notin $allowedUserIds) {
        Write-Log "Запрос от userId $userId отклонён - нет в разрешённых." -IsError
        return
    }

    # Получаем displayName из AD
    $displayName = Get-ADUserDisplayName -UserId $userId

    # Обработка команды /locked
    if ($text -eq "/locked" -or $text -eq "/locked@Unlock_locked_bot") {
        $lockedUsers = Get-LockedUsers
        if ($lockedUsers.Count -eq 0) {
            Send-TelegramMessage -ChatId $chatId -Text "Нет заблокированных учетных записей."
            Write-Log "Пользователь $displayName ($userId) запросил /locked. Заблокированных учетных записей не найдено."
        } else {
            $keyboard = Create-TelegramButtons -UserList $lockedUsers
            Send-TelegramMessage -ChatId $chatId -Text "Список заблокированных учетных записей:" -ReplyMarkup $keyboard
            Write-Log "Пользователь $displayName ($userId) запросил /locked. Найдено учетных записей: $($lockedUsers.Count)."
        }
    }

    # Обработка callback_query (нажатие на кнопку)
    if ($Update.PSObject.Properties["callback_query"] -and $Update.callback_query) {
        $callbackQuery = $Update.callback_query
        $buttonData = $callbackQuery.data

        if ([string]::IsNullOrEmpty($buttonData)) {
            Write-Log "Ошибка: callback_data пустое в callback_query от $userId." -IsError
            return
        }

        $result = Unlock-UserAccount -SamAccountName $buttonData
        Send-TelegramMessage -ChatId $chatId -Text $result -ThreadId $callbackQuery.message.message_thread_id
        Write-Log "Попытка разблокировки учетной записи $buttonData пользователем $displayName ($userId). Результат: $result"
    }
}

# Основной цикл проверки обновлений
$offset = 0
# Проверка возможности логирования при старте
if (-not (Ensure-LogDirectory) -or -not (Test-LogFile)) {
    Write-Host "[ERROR] $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') Не удалось инициализировать логирование в $logDir. Бот продолжит работу без записи логов." -ForegroundColor Red
}

while ($true) {
    try {
        # Проверка доступности Telegram API
        if (-not (Test-TelegramApi)) {
            Write-Log "Telegram API недоступен. Ожидание перед повторной попыткой..." -IsError
            Start-Sleep -Seconds $baseRetryDelay
            continue
        }

        $response = Invoke-RestMethod -Uri "$apiBaseUrl/getUpdates?offset=$offset&timeout=30" -Method Get -TimeoutSec $timeoutSeconds -ErrorAction Stop
        if ($response.result) {
            foreach ($update in $response.result) {
                if ($update.update_id -ge $offset) {
                    Handle-Update -Update $update
                    $offset = $update.update_id + 1
                } else {
                    Write-Log "Пропущено устаревшее обновление с update_id $($update.update_id)." -IsError
                }
            }
        }
    } catch {
        # Игнорируем ошибку таймаута
        if ($_.Exception.Message -notlike "*The request was canceled due to the configured HttpClient.Timeout*") {
            Write-Log "Ошибка при получении обновлений: $_" -IsError
        }
        $retryCount = 0
        while ($retryCount -lt $maxRetries) {
            $retryCount++
            $delay = $baseRetryDelay * [math]::Pow(2, $retryCount - 1)
            Start-Sleep -Seconds $delay
            try {
                $response = Invoke-RestMethod -Uri "$apiBaseUrl/getUpdates?offset=$offset&timeout=30" -Method Get -TimeoutSec $timeoutSeconds -ErrorAction Stop
                break
            } catch {
                if ($retryCount -eq $maxRetries -and $_.Exception.Message -notlike "*The request was canceled due to the configured HttpClient.Timeout*") {
                    Write-Log "Достигнуто максимальное количество попыток получения обновлений: $_" -IsError
                }
            }
        }
    }

    Start-Sleep -Milliseconds 500
}
