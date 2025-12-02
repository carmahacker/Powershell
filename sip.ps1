[Console]::OutputEncoding = [System.Text.Encoding]::GetEncoding("windows-1251")
Import-Module 'Posh-SSH'

# Данные Telegram API
$botToken = "7797221192:AAGIut7mcvFTUxBC3JhwscgX1111"
$apiURL = "https://api.telegram.org/bot$botToken"

# Список разрешённых пользователей
$allowedUserIds = @(
    287242111, 172918111, -1002173953111
)

# SSH данные
$username = 'aster-ip'
$keyfile  = "C:\infrabot\keys\id_ed25519"
$password = '1' | ConvertTo-SecureString -AsPlainText -Force
$credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $username, $password

# Функция поиска IP-адреса SIP
function Get-SIPInfo {
    param ([string]$sipNumber)

    if ($sipNumber -lt "1000" -or $sipNumber -gt "9999") {
        return "Введите команду в формате /sip XXXX (номер внутреннего телефона)"
    }

    # Определение сервера
    switch -Wildcard ($sipNumber) {
        '4*' { $sip_server = '10.72.225.225' }
        '8*' { $sip_server = '10.38.225.225' }
        '3*' { $sip_server = @('10.127.225.225', '10.126.225.225') }
        '7*' { $sip_server = @('10.79.225.225','10.78.225.225') }
        '2*' { $sip_server = '10.254.1.225' }
        '1*' { $sip_server = '10.1.225.225' }
        '5*' { $sip_server = '10.26.225.225' }
        '6*' { $sip_server = '10.25.225.225' }
        default { return "Неизвестный номер SIP." }
    }

    $command = "sudo /usr/sbin/asterisk -rx 'pjsip show endpoint $sipNumber' | grep 'Contact: ' | awk '{print $2}' | sed -E 's/.*@([^:]+):.*/\1/'"
    $rez_str = @()

    foreach ($sip_serv in $sip_server) {
        try {
            $SSession = New-SSHSession -HostName $sip_serv -Credential $credential -KeyFile $keyfile -AcceptKey
            $result = Invoke-SSHCommand -SSHSession $SSession -Command $command
            Remove-SSHSession $SSession

            # Фильтруем IP-адреса, убираем True
            $ip = ($result.Output | Select-String -Pattern "\b\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\b").Matches.Value
            if ($ip) {
                $rez_str += $ip
            }
        } catch {
            $rez_str += "Ошибка подключения к $sip_serv."
        }
    }

    # Проверка наличия данных
    if ($rez_str.Count -eq 0) {
        return "Такой номер телефона не зарегистрирован ни на одной АТС"
    }

    return ($rez_str -join "`n").Trim()
}

# Функция отправки сообщения
function Send-TelegramMessage {
    param ([string]$chatID, [string]$message)
    $parameters = @{ chat_id = $chatID; text = $message }
    Invoke-RestMethod -Uri "$apiURL/sendMessage" -Method Post -Body ($parameters | ConvertTo-Json -Depth 10) -ContentType "application/json"
}

# Основной цикл бота
$offset = 0
while ($true) {
    try {
        $updates = Invoke-RestMethod -Uri "$apiURL/getUpdates?offset=$offset" -Method Get

        foreach ($update in $updates.result) {
            $updateID = $update.update_id
            $chatID = $update.message.chat.id
            $userID = $update.message.from.id
            $messageText = $update.message.text

            if ($updateID -ge $offset) {
                $offset = $updateID + 1
            }

            # Проверка, разрешён ли пользователь
            if ($userID -notin $allowedUserIds) {
                Send-TelegramMessage -chatID $chatID -message "Доступ запрещён."
                continue
            }

            # Обработка команды
            if ($messageText -match "^/sip\s+(\d{4})$") {
                $sipNumber = $matches[1]
                $response = Get-SIPInfo -sipNumber $sipNumber
                Send-TelegramMessage -chatID $chatID -message $response
            }
        }
    } catch {
        Write-Host "Ошибка: $_"
    }
    Start-Sleep -Seconds 2
}
