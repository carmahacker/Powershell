<#
    HTTP JSON API Server with LDAP-based token authentication
    Features:
    - LDAP token validation using uSNChanged
    - UTF-8 (no BOM) JSON output
    - Dynamic /api/services (filters out system services)
#>

Add-Type -AssemblyName System.Net
Add-Type -AssemblyName System.IO
Add-Type -AssemblyName System.Security
Add-Type -AssemblyName System.DirectoryServices

# ----------------------------
# Настройки
# ----------------------------
$port = 10999
$ldapUser = "readldap@pharmasyntez.com"
$ldapAttr = "uSNChanged"
$ldapPath = "LDAP://dc3.pharmasyntez.com/DC=pharmasyntez,DC=com"

# ----------------------------
# LDAP: получение значения атрибута
# ----------------------------
function Get-LdapAttributeValue($userPrincipalName, $attribute) {
    try {
        $root = New-Object DirectoryServices.DirectoryEntry($ldapPath)
        $searcher = New-Object DirectoryServices.DirectorySearcher($root)
        $searcher.Filter = "(userPrincipalName=$userPrincipalName)"
        $searcher.PropertiesToLoad.Add($attribute) | Out-Null
        $result = $searcher.FindOne()
        if ($result) {
            return $result.Properties[$attribute][0]
        }
    } catch {
        Write-Warning "LDAP error: $_"
    }
    return $null
}

# ----------------------------
# Генерация токена, валидного 1 минуту
# ----------------------------
function Generate-Token($ldapValue) {
    if (-not $ldapValue) { return $null }
    $timeKey = (Get-Date -Format "yyyyMMddHHmm")
    $raw = "$ldapValue$timeKey"
    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($raw)
    $hash = $sha256.ComputeHash($bytes)
    return -join ($hash | ForEach-Object { $_.ToString("x2") })
}

# ----------------------------
# Проверка токена в запросе
# ----------------------------
function Check-Authorization($context) {
    try {
        $authHeader = $context.Request.Headers["Authorization"]
        if (-not $authHeader -or -not ($authHeader -match "^Bearer\s+(.+)$")) {
            return $false
        }

        $clientToken = $Matches[1].Trim()
        $ldapValue = Get-LdapAttributeValue $ldapUser $ldapAttr
        $serverToken = Generate-Token $ldapValue

        if ($clientToken -eq $serverToken) {
            return $true
        } else {
            Write-Host "❌ Invalid token from $($context.Request.RemoteEndPoint)"
            return $false
        }
    } catch {
        Write-Warning "Ошибка авторизации: $_"
        return $false
    }
}

# ----------------------------
# Отправка JSON-ответа (UTF-8 без BOM)
# ----------------------------
function Send-JsonResponse($context, $data, [int]$code = 200) {
    try {
        $json = $data | ConvertTo-Json -Depth 5
        $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
        $buffer = $utf8NoBom.GetBytes($json)

        $response = $context.Response
        $response.StatusCode = $code
        $response.ContentType = "application/json; charset=utf-8"
        $response.ContentEncoding = $utf8NoBom
        $response.ContentLength64 = $buffer.Length

        $response.OutputStream.Write($buffer, 0, $buffer.Length)
    } catch {
        Write-Warning "Ошибка отправки ответа: $_"
    } finally {
        try { $context.Response.OutputStream.Close() } catch {}
    }
}

# ----------------------------
# /api/status
# ----------------------------
function Get-SystemStatus {
    $cpu = Get-CimInstance Win32_Processor | Select-Object Name, NumberOfCores, NumberOfLogicalProcessors, MaxClockSpeed
    $ram = [math]::Round((Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory / 1GB, 2)
    $disks = Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3" |
        Select-Object DeviceID,
                      @{Name="SizeGB";Expression={[math]::Round($_.Size/1GB,2)}},
                      @{Name="FreeGB";Expression={[math]::Round($_.FreeSpace/1GB,2)}}
    $net = Get-NetIPAddress -AddressFamily IPv4 | Select-Object InterfaceAlias, IPAddress
    $dns = [System.Net.Dns]::GetHostEntry([System.Net.Dns]::GetHostName()).HostName
    $uptime = (Get-Date) - (Get-CimInstance Win32_OperatingSystem).LastBootUpTime

    return [PSCustomObject]@{
        CPU     = $cpu
        RAM_GB  = $ram
        Disks   = $disks
        Network = $net
        DNSName = $dns
        Uptime  = ("{0}d {1}h {2}m" -f $uptime.Days, $uptime.Hours, $uptime.Minutes)
    }
}

# ----------------------------
# /api/services — динамическое получение служб
# ----------------------------
function Get-ServicesStatus {
    try {
        $services = Get-CimInstance Win32_Service |
            Select-Object Name, DisplayName, Description, PathName, StartMode, State

        # Исключаем системные шаблоны
        $excludePatterns = @(
            'WinDefend', 'Defender', 'TrustedInstaller', 'WMPNetworkSvc',
            'Intel', 'ConfigMgr', 'SMS Agent', 'CcmExec', 'CmRcService', 'smstsmgr',
            'OSE', 'SUR QC', 'SystemUsageReportSvc', 'WdNisSvc',
            'Windows Media', 'Policy Platform', 'MDCoreSvc'
        )

        # Оставляем SQL и подобные
        $includeExceptions = @(
            'SQL', 'MSSQL', 'SQLSERVERAGENT', 'SQLBrowser', 'SQLWriter', 'SQLTELEMETRY'
        )

        $filtered = $services | Where-Object {
            $path = $_.PathName
            $name = $_.Name
            $display = $_.DisplayName

            $isExcluded = $false
            foreach ($p in $excludePatterns) {
                if ($name -match $p -or $display -match $p -or $path -match $p) {
                    $isExcluded = $true
                    break
                }
            }

            if (($path -match "C:\\Windows" -or $path -match "Microsoft") -and
                -not ($path -match "SQL Server")) {
                $isExcluded = $true
            }

            $isIncluded = $false
            foreach ($e in $includeExceptions) {
                if ($name -match $e -or $display -match $e -or $path -match $e) {
                    $isIncluded = $true
                    break
                }
            }

            (-not $isExcluded) -or $isIncluded
        }

        return $filtered
    } catch {
        Write-Warning "Ошибка при получении статусов служб: $_"
        return @{ error = "Internal error while reading services" }
    }
}

# ----------------------------
# /api/software
# ----------------------------
function Get-InstalledSoftware {
    Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*,
                     HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* |
    Where-Object { $_.DisplayName } |
    Select-Object DisplayName, DisplayVersion |
    Sort-Object DisplayName
}

# ----------------------------
# /api/firewall
# ----------------------------
function Get-FirewallListeners {
    Get-NetTCPConnection | Where-Object { $_.State -eq "Listen" } |
    ForEach-Object {
        $proc = Get-Process -Id $_.OwningProcess -ErrorAction SilentlyContinue
        [PSCustomObject]@{
            LocalAddress = $_.LocalAddress
            LocalPort    = $_.LocalPort
            ProcessName  = $proc.ProcessName
            PID          = $_.OwningProcess
        }
    } | Sort-Object LocalPort
}

# ----------------------------
# HTTP сервер
# ----------------------------
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://+:$port/")
$listener.Start()

Write-Host "✅ HTTP сервер запущен на порту $port"
Write-Host "Маршруты: /api/status, /api/services, /api/software, /api/firewall"
Write-Host "Требуется заголовок: Authorization: Bearer <token>"

try {
    while ($true) {
        $context = $listener.GetContext()
        $path = $context.Request.Url.AbsolutePath.ToLower()
        Write-Host "→ Запрос: $path"

        if ($path -eq "/favicon.ico") {
            Send-JsonResponse $context @{ message = "ignored" } 204
            continue
        }

        if (-not (Check-Authorization $context)) {
            Send-JsonResponse $context @{ error = "you are not authorized" } 401
            continue
        }

        switch ($path) {
            "/api/status"   { $response = Get-SystemStatus }
            "/api/services" { $response = Get-ServicesStatus }
            "/api/software" { $response = Get-InstalledSoftware }
            "/api/firewall" { $response = Get-FirewallListeners }
            default         { $response = @{ error = "Unknown endpoint: $path" } }
        }

        Send-JsonResponse $context $response
    }
}
catch {
    Write-Warning "Listener error: $_"
}
finally {
    if ($listener.IsListening) {
        $listener.Stop()
        $listener.Close()
    }
    Write-Host "Сервер остановлен."
}
