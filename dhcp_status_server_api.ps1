<#
.SYNOPSIS
  REST API для получения IP ↔ MAC (через ARP и DHCP).
  Поддерживает:
    • /api/status   — проверка состояния
    • /api/ping     — пинг хоста
    • /api/ip2mac   — поиск MAC по IP
    • /api/?mac=    — поиск IP по MAC
    • логирование запросов
    • ограничение по IP
#>

# === Настройки ===
$Port = 8080
$DhcpServer = $env:COMPUTERNAME
$AllowedClients = @("10.79.220.103", "10.79.77.3", "10.79.77.5", "10.254.1.32")
$LogFile = "C:\Scripts\Mac2IpApi.log"

# === DHCP модуль ===
try {
    Import-Module DhcpServer -ErrorAction Stop
    $dhcpAvailable = $true
} catch {
    $dhcpAvailable = $false
}

# === Утилиты ===
function Write-Log {
    param([string]$Message)
    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    Add-Content -Path $LogFile -Value "$timestamp | $Message"
}

function Normalize-Mac {
    param([string]$mac)
    $mac = $mac.ToLower()
    $mac = $mac -replace "([0-9a-f]{2})[^0-9a-f]?(?=.)", '$1:'
    return $mac.TrimEnd(':').ToUpper()
}

function Get-IpFromMac {
    param([string]$mac)
    $normalizedMac = Normalize-Mac $mac
    $ip = $null
    $source = "None"

    # ARP
    $arp = arp -a | Select-String $normalizedMac
    if ($arp) {
        $ip = ($arp -split '\s+')[1]
        $source = "ARP"
    }
    elseif ($dhcpAvailable) {
        try {
            $cleanMac = ($normalizedMac -replace '[:\-]', '').ToUpper()
            $scopes = Get-DhcpServerv4Scope -ComputerName $DhcpServer
            foreach ($s in $scopes) {
                $tmp = Get-DhcpServerv4Lease -ComputerName $DhcpServer -ScopeId $s.ScopeId |
                       Where-Object { ($_.ClientId -replace '[:\-]', '').ToUpper() -eq $cleanMac }
                if ($tmp) {
                    $ip = $tmp.IPAddress.ToString()
                    $source = "DHCP"
                    break
                }
            }
        } catch {
            Write-Warning "Ошибка при опросе DHCP: $_"
        }
    }
    return @($ip, $source, $normalizedMac)
}

function Get-MacFromIp {
    param([string]$ip)
    $mac = $null
    $source = "None"

    # ARP
    $arpLine = arp -a | Select-String " $ip "
    if ($arpLine) {
        $mac = ($arpLine -split '\s+')[2]
        $source = "ARP"
    }
    elseif ($dhcpAvailable) {
        try {
            $scopes = Get-DhcpServerv4Scope -ComputerName $DhcpServer
            foreach ($s in $scopes) {
                $lease = Get-DhcpServerv4Lease -ComputerName $DhcpServer -ScopeId $s.ScopeId |
                         Where-Object { $_.IPAddress -eq $ip }
                if ($lease) {
                    $mac = ($lease.ClientId -replace '([0-9A-Fa-f]{2})', '$1:').TrimEnd(':').ToUpper()
                    $source = "DHCP"
                    break
                }
            }
        } catch {
            Write-Warning "Ошибка при опросе DHCP: $_"
        }
    }
    return @($mac, $source)
}

function Test-HostReachable {
    param([string]$ip)
    try {
        $result = Test-Connection -ComputerName $ip -Count 1 -ErrorAction Stop
        return @{ reachable = $true; rtt = $result.ResponseTime }
    } catch {
        return @{ reachable = $false; rtt = $null }
    }
}

# === HTTP listener ===
Add-Type -AssemblyName System.Net
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://+:$Port/api/")
$listener.Start()
Write-Host "✅ MAC2IP API запущен на порту $Port..."
Write-Host "Разрешён доступ с: $($AllowedClients -join ', ')"
Write-Host "Лог: $LogFile"
Write-Host "Нажмите Ctrl+C для остановки."

# === Основной цикл ===
while ($listener.IsListening) {
    try {
        $context = $listener.GetContext()
        $request = $context.Request
        $response = $context.Response
        $remoteIP = $request.RemoteEndPoint.Address.ToString()
        $path = $request.Url.AbsolutePath.ToLower()

        # 🔒 Проверка доступа
        if ($AllowedClients -notcontains $remoteIP) {
            Write-Warning "⛔ Отказано в доступе: $remoteIP"
            $result = @{ status = "forbidden"; message = "Access denied for $remoteIP" }
            $json = $result | ConvertTo-Json
            $response.StatusCode = 403
            $buffer = [System.Text.Encoding]::UTF8.GetBytes($json)
            $response.OutputStream.Write($buffer, 0, $buffer.Length)
            $response.Close()
            continue
        }

        # --- /api/status ---
        if ($path -eq "/api/status") {
            try {
                $os = Get-CimInstance Win32_OperatingSystem
                $lastBoot = $os.LastBootUpTime
                $uptimeSpan = (Get-Date) - $lastBoot
                $uptime = "{0}d {1}h {2}m" -f $uptimeSpan.Days, $uptimeSpan.Hours, $uptimeSpan.Minutes
            } catch {
                $uptime = "unknown"
            }
            $result = @{
                status   = "running"
                hostname = $env:COMPUTERNAME
                uptime   = $uptime
                dhcp     = $dhcpAvailable
                logFile  = $LogFile
            }
        }

        # --- /api/ping ---
        elseif ($path -eq "/api/ping") {
            $ip = $request.QueryString["ip"]
            if (-not $ip) {
                $result = @{ status = "error"; message = "Parameter 'ip' is required" }
            } else {
                $ping = Test-HostReachable $ip
                $result = @{
                    status = "ok"
                    ip = $ip
                    reachable = $ping.reachable
                    rtt = $ping.rtt
                }
                Write-Log "$remoteIP | ping | $ip -> $($ping.reachable)"
            }
        }

        # --- /api/ip2mac ---
        elseif ($path -eq "/api/ip2mac") {
            $ip = $request.QueryString["ip"]
            if (-not $ip) {
                $result = @{ status = "error"; message = "Parameter 'ip' is required" }
            } else {
                $macInfo = Get-MacFromIp $ip
                $mac = $macInfo[0]
                $source = $macInfo[1]
                if ($mac) {
                    $result = @{
                        status = "ok"
                        ip = $ip
                        mac = $mac
                        source = $source
                    }
                } else {
                    $result = @{
                        status = "not_found"
                        ip = $ip
                        message = "MAC not found"
                    }
                }
                Write-Log "$remoteIP | ip2mac | $ip -> $mac ($source)"
            }
        }

        # --- /api/?mac= (MAC → IP) ---
        else {
            $mac = $request.QueryString["mac"]
            if (-not $mac) {
                $result = @{ status = "error"; message = "Parameter 'mac' is required" }
            } else {
                $res = Get-IpFromMac $mac
                $ip = $res[0]; $source = $res[1]; $normMac = $res[2]
                if ($ip) {
                    $ping = Test-HostReachable $ip
                    $result = @{
                        status = "ok"
                        mac = $normMac
                        ip = $ip
                        reachable = $ping.reachable
                        rtt = $ping.rtt
                        source = $source
                    }
                } else {
                    $result = @{
                        status = "not_found"
                        mac = $normMac
                        message = "IP not found in ARP or DHCP"
                    }
                }
                Write-Log "$remoteIP | mac2ip | $normMac -> $ip ($source)"
            }
        }

        # --- Ответ клиенту ---
        $json = $result | ConvertTo-Json -Depth 3
        $buffer = [System.Text.Encoding]::UTF8.GetBytes($json)
        $response.ContentType = "application/json"
        $response.OutputStream.Write($buffer, 0, $buffer.Length)
        $response.Close()
    } catch {
        Write-Log "Ошибка: $_"
        Write-Host "⚠ Ошибка: $_"
    }
}
